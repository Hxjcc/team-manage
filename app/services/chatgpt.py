"""
ChatGPT API 服务
用于调用 ChatGPT 后端 API,实现 Team 成员管理功能
"""
import asyncio
import logging
import time
from typing import Optional, Dict, Any, List
from curl_cffi.requests import AsyncSession
from app.services.settings import settings_service
from sqlalchemy.ext.asyncio import AsyncSession as DBAsyncSession

logger = logging.getLogger(__name__)


class ChatGPTService:
    """ChatGPT API 服务类"""

    BASE_URL = "https://chatgpt.com/backend-api"

    # 重试配置
    MAX_RETRIES = 3
    RETRY_DELAYS = [1, 2, 4]  # 指数退避: 1s, 2s, 4s

    # FlareSolverr CF cookies 缓存时间 (30 分钟)
    CF_COOKIE_TTL = 1800

    def __init__(self):
        """初始化 ChatGPT API 服务"""
        self.session: Optional[AsyncSession] = None
        self._cf_cookies: Optional[Dict[str, str]] = None
        self._cf_user_agent: Optional[str] = None
        self._cf_cookies_time: float = 0

    @staticmethod
    def _looks_like_html(text: Optional[str]) -> bool:
        if not text:
            return False
        stripped = text.lstrip().lower()
        return stripped.startswith("<!doctype html") or stripped.startswith("<html") or "<html" in stripped[:200]

    @staticmethod
    def _is_cloudflare_challenge(text: Optional[str]) -> bool:
        if not text:
            return False
        lowered = text.lower()
        return (
            "cdn-cgi/challenge-platform" in lowered
            or "_cf_chl_opt" in lowered
            or "cf-chl" in lowered
            or "enable javascript and cookies to continue" in lowered
        )

    @classmethod
    def _simplify_error_text(cls, text: Optional[str]) -> Dict[str, Optional[str]]:
        """
        将 HTML/超长错误收敛成可读提示，避免把整页 HTML 直接返回到前端。

        Returns:
            dict: { message: str, code: Optional[str] }
        """
        if not text:
            return {"message": "请求失败", "code": None}

        # Cloudflare 验证页（常见于 chatgpt.com 后端接口）
        if cls._looks_like_html(text) and cls._is_cloudflare_challenge(text):
            return {
                "message": "请求被 Cloudflare 拦截（需要浏览器验证）。请在系统设置中配置 FlareSolverr 后重试。",
                "code": "cloudflare_challenge",
            }

        # 其它 HTML 页面（如被重定向到登录页/错误页）
        if cls._looks_like_html(text):
            return {
                "message": "服务返回了 HTML 页面（可能被拦截或重定向），请稍后重试。",
                "code": "html_response",
            }

        # 普通文本：截断避免前端 toast 过长
        trimmed = str(text).strip()
        if len(trimmed) > 2000:
            trimmed = trimmed[:2000] + "...(已截断)"
        return {"message": trimmed, "code": None}

    async def _fetch_cf_cookies(self, db_session: DBAsyncSession) -> bool:
        """
        通过 FlareSolverr 获取 Cloudflare 验证 cookies

        Args:
            db_session: 数据库会话

        Returns:
            是否成功获取 cookies
        """
        config = await settings_service.get_flaresolverr_config(db_session)
        if not config["enabled"] or not config["url"]:
            return False

        flaresolverr_url = config["url"].rstrip("/") + "/v1"
        logger.info(f"通过 FlareSolverr 获取 CF cookies: {flaresolverr_url}")

        try:
            async with AsyncSession(timeout=120) as fs_session:
                response = await fs_session.post(
                    flaresolverr_url,
                    headers={"Content-Type": "application/json"},
                    json={
                        "cmd": "request.get",
                        "url": "https://chatgpt.com",
                        "maxTimeout": 60000
                    }
                )

                if response.status_code == 200:
                    data = response.json()
                    if data.get("status") == "ok":
                        solution = data.get("solution", {})
                        cookies_list = solution.get("cookies", [])
                        user_agent = solution.get("userAgent", "")

                        self._cf_cookies = {c["name"]: c["value"] for c in cookies_list}
                        self._cf_user_agent = user_agent or None
                        self._cf_cookies_time = time.time()

                        logger.info(f"FlareSolverr 成功: 获取 {len(self._cf_cookies)} 个 cookies")
                        return True
                    else:
                        logger.warning(f"FlareSolverr 失败: {data.get('message', '未知错误')}")
                else:
                    logger.warning(f"FlareSolverr HTTP 错误: {response.status_code}")

        except Exception as e:
            logger.error(f"FlareSolverr 异常: {e}")

        return False

    def _cf_cookies_valid(self) -> bool:
        """检查 CF cookies 缓存是否仍然有效"""
        return bool(self._cf_cookies) and (time.time() - self._cf_cookies_time) < self.CF_COOKIE_TTL

    async def _ensure_cf_cookies(self, db_session: DBAsyncSession):
        """确保有有效的 CF cookies (如果 FlareSolverr 已配置)"""
        if not self._cf_cookies_valid():
            await self._fetch_cf_cookies(db_session)

    async def _try_cf_recovery(self, db_session: DBAsyncSession) -> bool:
        """
        CF 验证失败后,尝试通过 FlareSolverr 重新获取 cookies 并重建会话

        Returns:
            是否恢复成功
        """
        logger.info("检测到 Cloudflare 验证,尝试通过 FlareSolverr 恢复...")

        # 清除缓存
        self._cf_cookies = None
        self._cf_user_agent = None
        self._cf_cookies_time = 0

        # 重新获取 cookies
        success = await self._fetch_cf_cookies(db_session)
        if success:
            # 重建 session
            if self.session:
                try:
                    await self.session.close()
                except Exception:
                    pass
                self.session = None
            self.session = await self._create_session(db_session)
            return True

        return False

    async def _create_session(self, db_session: DBAsyncSession) -> AsyncSession:
        """
        创建 HTTP 会话

        Args:
            db_session: 数据库会话

        Returns:
            curl_cffi AsyncSession 实例
        """
        # 如果 FlareSolverr 已配置,确保有 CF cookies
        await self._ensure_cf_cookies(db_session)

        # 创建会话 (使用 chrome 浏览器指纹)
        session = AsyncSession(
            impersonate="chrome",
            timeout=30
        )

        # 应用 CF cookies 到会话
        if self._cf_cookies:
            for name, value in self._cf_cookies.items():
                session.cookies.set(name, value, domain="chatgpt.com")
            logger.info(f"已应用 {len(self._cf_cookies)} 个 CF cookies 到会话")

        logger.info("创建 HTTP 会话")
        return session

    async def _make_request(
        self,
        method: str,
        url: str,
        headers: Dict[str, str],
        json_data: Optional[Dict[str, Any]] = None,
        db_session: Optional[DBAsyncSession] = None
    ) -> Dict[str, Any]:
        """
        发送 HTTP 请求 (带重试机制)

        Args:
            method: HTTP 方法 (GET/POST/DELETE)
            url: 请求 URL
            headers: 请求头
            json_data: JSON 请求体
            db_session: 数据库会话

        Returns:
            响应数据字典,包含 success, status_code, data, error
        """
        # 创建会话
        if not self.session:
            self.session = await self._create_session(db_session)

        cf_retried = False  # 标记是否已通过 FlareSolverr 重试过

        # 重试循环
        for attempt in range(self.MAX_RETRIES):
            try:
                logger.info(f"发送请求: {method} {url} (尝试 {attempt + 1}/{self.MAX_RETRIES})")

                # 如果有 FlareSolverr 的 User-Agent,覆盖请求头
                request_headers = dict(headers)
                if self._cf_user_agent:
                    request_headers["User-Agent"] = self._cf_user_agent

                # 发送请求
                if method == "GET":
                    response = await self.session.get(url, headers=request_headers)
                elif method == "POST":
                    response = await self.session.post(url, headers=request_headers, json=json_data)
                elif method == "DELETE":
                    response = await self.session.delete(url, headers=request_headers, json=json_data)
                else:
                    raise ValueError(f"不支持的 HTTP 方法: {method}")

                status_code = response.status_code
                logger.info(f"响应状态码: {status_code}")

                # 2xx 成功
                if 200 <= status_code < 300:
                    # 若返回 HTML（Cloudflare/重定向页面），即便是 2xx 也应视为失败
                    content_type = ""
                    try:
                        content_type = (response.headers.get("content-type") or "").lower()
                    except Exception:
                        content_type = ""

                    is_json = "application/json" in content_type
                    text_body = None

                    if is_json:
                        try:
                            data = response.json()
                            return {
                                "success": True,
                                "status_code": status_code,
                                "data": data,
                                "error": None
                            }
                        except Exception:
                            text_body = response.text
                    else:
                        # 非 JSON 情况下尝试解析；若失败或内容像 HTML，则报错
                        try:
                            data = response.json()
                            return {
                                "success": True,
                                "status_code": status_code,
                                "data": data,
                                "error": None
                            }
                        except Exception:
                            text_body = response.text

                    simplified = self._simplify_error_text(text_body)

                    # CF 验证检测: 尝试通过 FlareSolverr 恢复
                    if simplified.get("code") == "cloudflare_challenge" and not cf_retried and db_session:
                        recovery_ok = await self._try_cf_recovery(db_session)
                        if recovery_ok:
                            cf_retried = True
                            continue

                    logger.warning(f"响应内容非 JSON 或解析失败: {simplified['message']}")
                    return {
                        "success": False,
                        "status_code": status_code,
                        "data": None,
                        "error": simplified["message"],
                        "error_code": simplified.get("code") or "invalid_response"
                    }

                # 4xx 客户端错误 (不重试)
                if 400 <= status_code < 500:
                    error_code = None
                    try:
                        error_data = response.json()
                        error_msg = error_data.get("detail", response.text)

                        # 检测特定错误码
                        if isinstance(error_data, dict):
                            # 有些错误可能在 error 字段里
                            error_info = error_data.get("error")
                            if isinstance(error_info, dict):
                                error_code = error_info.get("code")
                            else:
                                error_code = error_data.get("code")
                    except Exception:
                        error_msg = response.text

                    simplified = self._simplify_error_text(error_msg)
                    error_msg = simplified["message"]
                    error_code = error_code or simplified.get("code")

                    # 4xx 也可能是 CF 挑战 (403)
                    if simplified.get("code") == "cloudflare_challenge" and not cf_retried and db_session:
                        recovery_ok = await self._try_cf_recovery(db_session)
                        if recovery_ok:
                            cf_retried = True
                            continue

                    logger.warning(f"客户端错误 {status_code}: {error_msg} (code: {error_code})")

                    return {
                        "success": False,
                        "status_code": status_code,
                        "data": None,
                        "error": error_msg,
                        "error_code": error_code
                    }

                # 5xx 服务器错误 (需要重试)
                if status_code >= 500:
                    # Cloudflare 验证页有时会以 5xx 返回
                    try:
                        body_text = response.text
                    except Exception:
                        body_text = ""

                    if self._looks_like_html(body_text) and self._is_cloudflare_challenge(body_text):
                        # 尝试通过 FlareSolverr 恢复
                        if not cf_retried and db_session:
                            recovery_ok = await self._try_cf_recovery(db_session)
                            if recovery_ok:
                                cf_retried = True
                                continue

                        simplified = self._simplify_error_text(body_text)
                        logger.warning(f"服务器错误 {status_code}: {simplified['message']}")
                        return {
                            "success": False,
                            "status_code": status_code,
                            "data": None,
                            "error": simplified["message"],
                            "error_code": simplified.get("code") or "cloudflare_challenge"
                        }

                    logger.warning(f"服务器错误 {status_code},准备重试")

                    # 如果不是最后一次尝试,等待后重试
                    if attempt < self.MAX_RETRIES - 1:
                        delay = self.RETRY_DELAYS[attempt]
                        logger.info(f"等待 {delay}s 后重试")
                        await asyncio.sleep(delay)
                        continue

                    # 最后一次尝试失败
                    return {
                        "success": False,
                        "status_code": status_code,
                        "data": None,
                        "error": f"服务器错误 {status_code},已重试 {self.MAX_RETRIES} 次"
                    }

            except asyncio.TimeoutError:
                logger.warning(f"请求超时 (尝试 {attempt + 1}/{self.MAX_RETRIES})")

                # 如果不是最后一次尝试,等待后重试
                if attempt < self.MAX_RETRIES - 1:
                    delay = self.RETRY_DELAYS[attempt]
                    logger.info(f"等待 {delay}s 后重试")
                    await asyncio.sleep(delay)
                    continue

                # 最后一次尝试失败
                return {
                    "success": False,
                    "status_code": 0,
                    "data": None,
                    "error": f"请求超时,已重试 {self.MAX_RETRIES} 次"
                }

            except Exception as e:
                logger.error(f"请求异常: {e}")

                # 如果不是最后一次尝试,等待后重试
                if attempt < self.MAX_RETRIES - 1:
                    delay = self.RETRY_DELAYS[attempt]
                    logger.info(f"等待 {delay}s 后重试")
                    await asyncio.sleep(delay)
                    continue

                # 最后一次尝试失败
                return {
                    "success": False,
                    "status_code": 0,
                    "data": None,
                    "error": f"请求异常: {str(e)}"
                }

        # 不应该到达这里
        return {
            "success": False,
            "status_code": 0,
            "data": None,
            "error": "未知错误"
        }

    async def send_invite(
        self,
        access_token: str,
        account_id: str,
        email: str,
        db_session: DBAsyncSession
    ) -> Dict[str, Any]:
        """
        发送 Team 邀请

        Args:
            access_token: AT Token
            account_id: Account ID
            email: 邀请的邮箱地址
            db_session: 数据库会话

        Returns:
            结果字典,包含 success, status_code, error
        """
        url = f"{self.BASE_URL}/accounts/{account_id}/invites"

        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {access_token}",
            "chatgpt-account-id": account_id
        }

        json_data = {
            "email_addresses": [email],
            "role": "standard-user",
            "resend_emails": True
        }

        logger.info(f"发送邀请: {email} -> Team {account_id}")

        result = await self._make_request("POST", url, headers, json_data, db_session)

        # 特殊处理 409 (用户已是成员)
        if result["status_code"] == 409:
            result["error"] = "用户已是该 Team 的成员"

        # 特殊处理 422 (Team 已满或邮箱格式错误)
        if result["status_code"] == 422:
            result["error"] = "Team 已满或邮箱格式错误"

        return result

    async def get_members(
        self,
        access_token: str,
        account_id: str,
        db_session: DBAsyncSession
    ) -> Dict[str, Any]:
        """
        获取 Team 成员列表

        Args:
            access_token: AT Token
            account_id: Account ID
            db_session: 数据库会话

        Returns:
            结果字典,包含 success, members (成员列表), total (总数), error
        """
        all_members = []
        offset = 0
        limit = 50

        while True:
            url = f"{self.BASE_URL}/accounts/{account_id}/users?limit={limit}&offset={offset}"

            headers = {
                "Authorization": f"Bearer {access_token}"
            }

            logger.info(f"获取成员列表: Team {account_id}, offset={offset}")

            result = await self._make_request("GET", url, headers, db_session=db_session)

            if not result["success"]:
                return {
                    "success": False,
                    "members": [],
                    "total": 0,
                    "error": result["error"],
                    "error_code": result.get("error_code")
                }

            # 解析响应
            data = result["data"]
            items = data.get("items", [])
            total = data.get("total", 0)

            all_members.extend(items)

            # 检查是否还有更多成员
            if len(all_members) >= total:
                break

            offset += limit

        logger.info(f"获取成员列表成功: 共 {len(all_members)} 个成员")

        return {
            "success": True,
            "members": all_members,
            "total": len(all_members),
            "error": None
        }

    async def get_invites(
        self,
        access_token: str,
        account_id: str,
        db_session: DBAsyncSession
    ) -> Dict[str, Any]:
        """
        获取 Team 待加入成员列表 (邀请列表)

        Args:
            access_token: AT Token
            account_id: Account ID
            db_session: 数据库会话

        Returns:
            结果字典,包含 success, items (邀请列表), total (总数), error
        """
        url = f"{self.BASE_URL}/accounts/{account_id}/invites"

        headers = {
            "Authorization": f"Bearer {access_token}",
            "chatgpt-account-id": account_id
        }

        logger.info(f"获取邀请列表: Team {account_id}")

        result = await self._make_request("GET", url, headers, db_session=db_session)

        if not result["success"]:
            return {
                "success": False,
                "items": [],
                "total": 0,
                "error": result["error"],
                "error_code": result.get("error_code")
            }

        data = result["data"]
        items = data.get("items", [])
        total = data.get("total", len(items))

        return {
            "success": True,
            "items": items,
            "total": total,
            "error": None
        }

    async def delete_invite(
        self,
        access_token: str,
        account_id: str,
        email: str,
        db_session: DBAsyncSession
    ) -> Dict[str, Any]:
        """
        撤回 Team 邀请

        Args:
            access_token: AT Token
            account_id: Account ID
            email: 邀请的邮箱地址
            db_session: 数据库会话

        Returns:
            结果字典,包含 success, status_code, error
        """
        url = f"{self.BASE_URL}/accounts/{account_id}/invites"

        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {access_token}",
            "chatgpt-account-id": account_id
        }

        json_data = {
            "email_address": email
        }

        logger.info(f"撤回邀请: {email} from Team {account_id}")

        result = await self._make_request("DELETE", url, headers, json_data, db_session)

        return result

    async def delete_member(
        self,
        access_token: str,
        account_id: str,
        user_id: str,
        db_session: DBAsyncSession
    ) -> Dict[str, Any]:
        """
        删除 Team 成员

        Args:
            access_token: AT Token
            account_id: Account ID
            user_id: 用户 ID (格式: user-xxx)
            db_session: 数据库会话

        Returns:
            结果字典,包含 success, status_code, error
        """
        url = f"{self.BASE_URL}/accounts/{account_id}/users/{user_id}"

        headers = {
            "Authorization": f"Bearer {access_token}",
            "chatgpt-account-id": account_id
        }

        logger.info(f"删除成员: {user_id} from Team {account_id}")

        result = await self._make_request("DELETE", url, headers, db_session=db_session)

        # 特殊处理 403 (无权限删除 owner)
        if result["status_code"] == 403:
            result["error"] = "无权限删除该成员 (可能是 owner)"

        # 特殊处理 404 (用户不存在)
        if result["status_code"] == 404:
            result["error"] = "用户不存在"

        return result

    async def get_account_info(
        self,
        access_token: str,
        db_session: DBAsyncSession
    ) -> Dict[str, Any]:
        """
        获取 account-id 和订阅信息

        Args:
            access_token: AT Token
            db_session: 数据库会话

        Returns:
            结果字典,包含 success, accounts (账户列表), error
        """
        url = f"{self.BASE_URL}/accounts/check/v4-2023-04-27"

        headers = {
            "Authorization": f"Bearer {access_token}"
        }

        logger.info("获取 account-id 和订阅信息")

        result = await self._make_request("GET", url, headers, db_session=db_session)

        if not result["success"]:
            return {
                "success": False,
                "accounts": [],
                "error": result["error"],
                "error_code": result.get("error_code")
            }

        # 解析响应
        data = result["data"]
        accounts_data = data.get("accounts", {})

        # 提取所有 Team 类型的账户
        team_accounts = []
        for account_id, account_info in accounts_data.items():
            account = account_info.get("account", {})
            entitlement = account_info.get("entitlement", {})

            # 只保留 Team 类型的账户
            if account.get("plan_type") == "team":
                team_accounts.append({
                    "account_id": account_id,
                    "name": account.get("name", ""),
                    "plan_type": account.get("plan_type", ""),
                    "subscription_plan": entitlement.get("subscription_plan", ""),
                    "expires_at": entitlement.get("expires_at", ""),
                    "has_active_subscription": entitlement.get("has_active_subscription", False)
                })

        logger.info(f"获取账户信息成功: 共 {len(team_accounts)} 个 Team 账户")

        return {
            "success": True,
            "accounts": team_accounts,
            "error": None
        }

    async def refresh_access_token_with_session_token(
        self,
        session_token: str,
        db_session: DBAsyncSession
    ) -> Dict[str, Any]:
        """
        使用 session_token 刷新 access_token
        
        Args:
            session_token: session_token
            db_session: 数据库会话
            
        Returns:
            结果字典,包含 success, access_token, error
        """
        url = "https://chatgpt.com/api/auth/session"
        
        headers = {
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
            "User-Agent": self._cf_user_agent or "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        }

        cookies = {
            "__Secure-next-auth.session-token": session_token
        }

        logger.info("使用 session_token 刷新 access_token")

        if not self.session:
            self.session = await self._create_session(db_session)

        try:
            response = await self.session.get(url, headers=headers, cookies=cookies)
            status_code = response.status_code
            if status_code == 200:
                data = response.json()
                access_token = data.get("accessToken")
                if access_token:
                    return {
                        "success": True,
                        "access_token": access_token
                    }
                return {"success": False, "error": "响应中未包含 accessToken"}
            else:
                error_code = None
                error_msg = response.text
                try:
                    error_data = response.json()
                    error_msg = error_data.get("detail", error_msg)
                    if isinstance(error_data, dict):
                        error_info = error_data.get("error")
                        if isinstance(error_info, dict):
                            error_code = error_info.get("code")
                        else:
                            error_code = error_data.get("code")
                except Exception:
                    pass
                
                logger.warning(f"session_token 刷新失败 {status_code}: {error_msg} (code: {error_code})")
                return {
                    "success": False, 
                    "status_code": status_code,
                    "error": error_msg,
                    "error_code": error_code
                }
        except Exception as e:
            logger.error(f"session_token 刷新失败: {e}")
            return {"success": False, "error": str(e)}

    async def refresh_access_token_with_refresh_token(
        self,
        refresh_token: str,
        client_id: str,
        db_session: DBAsyncSession
    ) -> Dict[str, Any]:
        """
        使用 refresh_token 刷新 access_token
        
        Args:
            refresh_token: refresh_token
            client_id: client_id
            db_session: 数据库会话
            
        Returns:
            结果字典,包含 success, access_token, refresh_token, error
        """
        url = "https://auth.openai.com/oauth/token"
        
        json_data = {
            "client_id": client_id,
            "grant_type": "refresh_token",
            "redirect_uri": "com.openai.sora://auth.openai.com/android/com.openai.sora/callback",
            "refresh_token": refresh_token
        }
        
        headers = {
            "Content-Type": "application/json",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        }
        
        logger.info("使用 refresh_token 刷新 access_token")
        
        if not self.session:
            self.session = await self._create_session(db_session)
            
        try:
            response = await self.session.post(url, headers=headers, json=json_data)
            status_code = response.status_code
            if status_code == 200:
                data = response.json()
                return {
                    "success": True,
                    "access_token": data.get("access_token"),
                    "refresh_token": data.get("refresh_token")
                }
            else:
                error_code = None
                error_msg = response.text
                try:
                    error_data = response.json()
                    # OAuth 错误通常在 'error' 字段(字符串)中, 详细在 'error_description'
                    if isinstance(error_data, dict):
                        error_code = error_data.get("error")
                        error_msg = error_data.get("error_description", error_msg)
                except Exception:
                    pass

                logger.warning(f"refresh_token 刷新失败 {status_code}: {error_msg} (code: {error_code})")
                return {
                    "success": False,
                    "status_code": status_code,
                    "error": error_msg,
                    "error_code": error_code
                }
        except Exception as e:
            logger.error(f"refresh_token 刷新失败: {e}")
            return {"success": False, "error": str(e)}

    async def close(self):
        """关闭 HTTP 会话"""
        if self.session:
            await self.session.close()
            self.session = None
            logger.info("HTTP 会话已关闭")

    async def clear_session(self):
        """清理当前会话和 CF cookies 缓存"""
        self._cf_cookies = None
        self._cf_user_agent = None
        self._cf_cookies_time = 0
        await self.close()


# 创建全局实例
chatgpt_service = ChatGPTService()
