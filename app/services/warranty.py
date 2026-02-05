"""
质保服务
处理用户质保查询和验证
"""
import logging
from typing import Optional, Dict, Any, List
from datetime import datetime, timedelta
from sqlalchemy import select, and_, or_, func
from sqlalchemy.ext.asyncio import AsyncSession

from app.models import RedemptionCode, RedemptionRecord, Team
from app.utils.time_utils import get_now

logger = logging.getLogger(__name__)

# 全局频率限制字典: {(type, key): last_time}
# type: 'email' or 'code'
_query_rate_limit = {}


class WarrantyService:
    """质保服务类"""

    def __init__(self):
        """初始化质保服务"""
        from app.services.team import TeamService
        self.team_service = TeamService()

    async def check_warranty_status(
        self,
        db_session: AsyncSession,
        email: Optional[str] = None,
        code: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        检查用户质保状态

        Args:
            db_session: 数据库会话
            email: 用户邮箱
            code: 兑换码

        Returns:
            结果字典,包含 success, has_warranty, warranty_valid, warranty_expires_at, 
            banned_teams, can_reuse, original_code, error
        """
        try:
            if not email and not code:
                return {
                    "success": False,
                    "error": "必须提供邮箱或兑换码"
                }

            # 0. 频率限制 (每个邮箱或每个码 30 秒只能查一次)
            now = datetime.now()
            limit_key = ("email", email) if email else ("code", code)
            last_time = _query_rate_limit.get(limit_key)
            if last_time and (now - last_time).total_seconds() < 30:
                wait_time = int(30 - (now - last_time).total_seconds())
                return {
                    "success": False,
                    "error": f"查询太频繁,请 {wait_time} 秒后再试"
                }
            _query_rate_limit[limit_key] = now

            # 1. 查找兑换记录和相关联的 Team, Code
            records_data = []

            if code:
                # 通过兑换码查找所有关联记录
                stmt = (
                    select(RedemptionRecord, RedemptionCode, Team)
                    .join(RedemptionCode, RedemptionRecord.code == RedemptionCode.code)
                    .join(Team, RedemptionRecord.team_id == Team.id)
                    .where(RedemptionCode.code == code)
                    .order_by(RedemptionRecord.redeemed_at.desc())
                )
                result = await db_session.execute(stmt)
                first_record = result.first()
                if first_record:
                    records_data = [first_record]
                else:
                    records_data = []

                # 如果没有记录，可能是码还没被使用或不存在
                if not records_data:
                    stmt = select(RedemptionCode).where(RedemptionCode.code == code)
                    result = await db_session.execute(stmt)
                    redemption_code_obj = result.scalar_one_or_none()
                    
                    if not redemption_code_obj:
                        return {
                            "success": True,
                            "has_warranty": False,
                            "warranty_valid": False,
                            "warranty_expires_at": None,
                            "banned_teams": [],
                            "can_reuse": False,
                            "original_code": None,
                            "records": [],
                            "message": "兑换码不存在"
                        }
                    
                    # 只有码没有记录的情况
                    # 质保默认随绑定 Team 到期（Team 到期质保即结束）
                    display_expiry = None
                    if redemption_code_obj.has_warranty and redemption_code_obj.bound_team_id:
                        stmt = select(Team).where(Team.id == redemption_code_obj.bound_team_id)
                        team_result = await db_session.execute(stmt)
                        bound_team = team_result.scalar_one_or_none()
                        if bound_team and bound_team.expires_at:
                            display_expiry = bound_team.expires_at

                    if not display_expiry and redemption_code_obj.has_warranty and redemption_code_obj.status != "unused":
                        display_expiry = redemption_code_obj.warranty_expires_at

                    warranty_valid = False
                    if redemption_code_obj.has_warranty:
                        if display_expiry:
                            warranty_valid = display_expiry > get_now()
                        else:
                            # 无法推断到期时间(通常不会发生)，兜底视为有效
                            warranty_valid = True

                    return {
                        "success": True,
                        "has_warranty": redemption_code_obj.has_warranty,
                        "warranty_valid": warranty_valid,
                        "warranty_expires_at": display_expiry.isoformat() if display_expiry else None,
                        "banned_teams": [],
                        "can_reuse": False,
                        "original_code": redemption_code_obj.code,
                        "records": [{
                            "code": redemption_code_obj.code,
                            "has_warranty": redemption_code_obj.has_warranty,
                            "warranty_valid": warranty_valid,
                            "status": redemption_code_obj.status,
                            "used_at": None,
                            "team_id": None,
                            "team_name": None,
                            "team_status": None,
                            "team_expires_at": None,
                            "warranty_expires_at": display_expiry.isoformat() if display_expiry else None
                        }],
                        "message": "兑换码尚未被使用"
                    }

            elif email:
                # 通过邮箱查找所有兑换记录
                stmt = (
                    select(RedemptionRecord, RedemptionCode, Team)
                    .join(RedemptionCode, RedemptionRecord.code == RedemptionCode.code)
                    .join(Team, RedemptionRecord.team_id == Team.id)
                    .where(RedemptionRecord.email == email)
                    .order_by(RedemptionRecord.redeemed_at.desc())
                )
                result = await db_session.execute(stmt)
                all_records = result.all()

                # 只保留每个兑换码的最近一条记录
                seen_codes = set()
                records_data = []
                for row in all_records:
                    # row format: (RedemptionRecord, RedemptionCode, Team)
                    record_obj = row[0]
                    if record_obj.code not in seen_codes:
                        seen_codes.add(record_obj.code)
                        records_data.append(row)

            if not records_data:
                return {
                    "success": True,
                    "has_warranty": False,
                    "warranty_valid": False,
                    "warranty_expires_at": None,
                    "banned_teams": [],
                    "can_reuse": False,
                    "original_code": None,
                    "records": [],
                    "message": "未找到兑换记录"
                }

            # 2. 处理记录并进行必要的实时同步
            final_records = []
            banned_teams_info = []
            has_any_warranty = False
            primary_warranty_valid = False
            primary_expiry = None
            primary_code = None
            can_reuse = False

            # 为动态兜底计算准备“首次使用时间”(min redeemed_at)
            code_set = {code_obj.code for _, code_obj, _ in records_data if code_obj.has_warranty}
            activation_map = {}
            if code_set:
                stmt = (
                    select(RedemptionRecord.code, func.min(RedemptionRecord.redeemed_at))
                    .where(RedemptionRecord.code.in_(list(code_set)))
                    .group_by(RedemptionRecord.code)
                )
                result = await db_session.execute(stmt)
                activation_map = {c: activated_at for c, activated_at in result.all()}

            # 预取绑定 Team (用于按绑定 Team 到期时间计算质保)
            bound_team_ids = {
                int(code_obj.bound_team_id)
                for _, code_obj, _ in records_data
                if code_obj.has_warranty and code_obj.bound_team_id
            }
            bound_team_map = {}
            if bound_team_ids:
                stmt = select(Team).where(Team.id.in_(list(bound_team_ids)))
                result = await db_session.execute(stmt)
                bound_team_map = {int(t.id): t for t in result.scalars().all()}

            for record, code_obj, team in records_data:
                # 同步 Team 状态
                if team.status not in ["banned", "error"]:
                    logger.info(f"质保查询: 正在实时测试 Team {team.id} ({team.team_name}) 的状态")
                    await self.team_service.sync_team_info(team.id, db_session)
                    # 同步后 team 对象的属性会自动更新

                # 动态计算/提取质保信息
                expiry_date = None
                if code_obj.has_warranty:
                    effective_team = bound_team_map.get(int(code_obj.bound_team_id)) if code_obj.bound_team_id else None
                    if effective_team and effective_team.expires_at:
                        expiry_date = effective_team.expires_at
                    elif team.expires_at:
                        expiry_date = team.expires_at
                    elif code_obj.warranty_expires_at:
                        expiry_date = code_obj.warranty_expires_at
                    else:
                        # 兜底：首次使用时间 + 质保天数
                        start_time = activation_map.get(code_obj.code) or code_obj.used_at or record.redeemed_at
                        if start_time:
                            days = code_obj.warranty_days or 30
                            expiry_date = start_time + timedelta(days=days)

                is_valid = False
                if code_obj.has_warranty:
                    is_valid = True if expiry_date and expiry_date > get_now() else False

                if code_obj.has_warranty:
                    has_any_warranty = True
                    # 以最近的一个质保码作为主要质保状态参考
                    if primary_code is None:
                        primary_warranty_valid = is_valid
                        primary_expiry = expiry_date
                        primary_code = code_obj.code

                # 记录封号 Team
                if team.status == "banned":
                    banned_teams_info.append({
                        "team_id": team.id,
                        "team_name": team.team_name,
                        "email": team.email,
                        "banned_at": team.last_sync.isoformat() if team.last_sync else None
                    })

                final_records.append({
                    "code": code_obj.code,
                    "has_warranty": code_obj.has_warranty,
                    "warranty_valid": is_valid,
                    "warranty_expires_at": expiry_date.isoformat() if expiry_date else None,
                    "status": code_obj.status,
                    "used_at": record.redeemed_at.isoformat() if record.redeemed_at else None,
                    "team_id": team.id,
                    "team_name": team.team_name,
                    "team_status": team.status,
                    "team_expires_at": team.expires_at.isoformat() if team.expires_at else None,
                    "email": record.email
                })

            # 3. 判断是否可以重复使用 (只要有有效的质保码且有被封的 Team)
            if has_any_warranty and primary_warranty_valid and len(banned_teams_info) > 0:
                # 进一步验证 (使用现有的 validate_warranty_reuse 逻辑)
                # 这里为了简单直接复用逻辑判断
                can_reuse = True

            return {
                "success": True,
                "has_warranty": has_any_warranty,
                "warranty_valid": primary_warranty_valid,
                "warranty_expires_at": primary_expiry.isoformat() if primary_expiry else None,
                "banned_teams": banned_teams_info,
                "can_reuse": can_reuse,
                "original_code": primary_code,
                "records": final_records,
                "message": "查询成功"
            }

        except Exception as e:
            logger.error(f"检查质保状态失败: {e}")
            return {
                "success": False,
                "error": f"检查质保状态失败: {str(e)}"
            }

    async def validate_warranty_reuse(
        self,
        db_session: AsyncSession,
        code: str,
        email: str
    ) -> Dict[str, Any]:
        """
        验证质保码是否可重复使用

        Args:
            db_session: 数据库会话
            code: 兑换码
            email: 用户邮箱

        Returns:
            结果字典,包含 success, can_reuse, reason, error
        """
        try:
            # 1. 查询兑换码
            stmt = select(RedemptionCode).where(RedemptionCode.code == code)
            result = await db_session.execute(stmt)
            redemption_code = result.scalar_one_or_none()

            if not redemption_code:
                return {
                    "success": True,
                    "can_reuse": False,
                    "reason": "兑换码不存在",
                    "error": None
                }

            # 2. 检查是否为质保码
            if not redemption_code.has_warranty:
                return {
                    "success": True,
                    "can_reuse": False,
                    "reason": "该兑换码不是质保兑换码",
                    "error": None
                }

            # 3. 检查质保期是否有效
            now = get_now()
            expiry_dt = None

            # 优先取绑定 Team 的到期时间（Team 到期质保即结束）
            if redemption_code.bound_team_id:
                stmt = select(Team).where(Team.id == redemption_code.bound_team_id)
                team_result = await db_session.execute(stmt)
                bound_team = team_result.scalar_one_or_none()
                if bound_team and bound_team.expires_at:
                    expiry_dt = bound_team.expires_at

            # 没有绑定 Team 时，尝试取首次加入的 Team 到期时间
            if not expiry_dt:
                stmt = (
                    select(Team.expires_at)
                    .join(RedemptionRecord, RedemptionRecord.team_id == Team.id)
                    .where(RedemptionRecord.code == code)
                    .order_by(RedemptionRecord.redeemed_at.asc())
                    .limit(1)
                )
                result = await db_session.execute(stmt)
                first_team_expires_at = result.scalar_one_or_none()
                if first_team_expires_at:
                    expiry_dt = first_team_expires_at

            # 兜底：使用已写入的质保到期时间
            if not expiry_dt and redemption_code.warranty_expires_at:
                expiry_dt = redemption_code.warranty_expires_at

            # 最终兜底：首次使用时间 + 质保天数
            if not expiry_dt:
                stmt = select(func.min(RedemptionRecord.redeemed_at)).where(RedemptionRecord.code == code)
                result = await db_session.execute(stmt)
                activated_at = result.scalar_one_or_none()
                days = redemption_code.warranty_days or 30
                if activated_at:
                    expiry_dt = activated_at + timedelta(days=days)
                elif redemption_code.used_at:
                    expiry_dt = redemption_code.used_at + timedelta(days=days)

            if expiry_dt and expiry_dt < now:
                return {
                    "success": True,
                    "can_reuse": False,
                    "reason": "质保已过期",
                    "error": None
                }

            if not expiry_dt:
                return {
                    "success": True,
                    "can_reuse": False,
                    "reason": "质保状态异常，请联系管理员",
                    "error": None
                }

            # 4. 检查该兑换码当前是否已有正在使用的活跃 Team (全局检查，不限邮箱)
            # 逻辑：如果该码名下有任何一个 Team 还是 active/full 状态且未过期，则不允许新的激活
            stmt = select(RedemptionRecord).where(RedemptionRecord.code == code)
            result = await db_session.execute(stmt)
            all_records_for_code = result.scalars().all()
            
            for record in all_records_for_code:
                stmt = select(Team).where(Team.id == record.team_id)
                result = await db_session.execute(stmt)
                team = result.scalar_one_or_none()
                
                if team:
                    is_expired = team.expires_at and team.expires_at < get_now()
                    if team.status in ["active", "full"] and not is_expired:
                        # 如果是同一个邮箱，提示已在有效 Team 中
                        if record.email == email:
                            return {
                                "success": True,
                                "can_reuse": False,
                                "reason": f"您已在有效 Team 中 ({team.team_name or team.id})，不可重复兑换",
                                "error": None
                            }
                        else:
                            # 如果是不同邮箱，提示已被占用
                            return {
                                "success": True,
                                "can_reuse": False,
                                "reason": "该兑换码当前已被其他账号绑定且正在使用中。如需更换，请确保原账号已下车或原 Team 已失效。",
                                "error": None
                            }

            # 5. 查找当前用户使用该兑换码的记录 (用于后续逻辑判断)
            records = [r for r in all_records_for_code if r.email == email]
            
            if not records:
                # 之前没有该邮箱的记录，但上面已经检查过没有其他活跃 Team 了，所以允许“新开”或“接手”
                return {
                    "success": True,
                    "can_reuse": True,
                    "reason": "可更名使用 (或首次使用)",
                    "error": None
                }

            # 5. 检查用户当前是否已在有效的 Team 中
            # 逻辑：如果最近一次加入的 Team 仍然有效（active/full 且未过期），则不允许重复使用
            for record in records:
                stmt = select(Team).where(Team.id == record.team_id)
                result = await db_session.execute(stmt)
                team = result.scalar_one_or_none()
                
                if team:
                    # 如果有任何一个关联 Team 还是 active/full 状态，且未过期
                    is_expired = team.expires_at and team.expires_at < get_now()
                    if team.status in ["active", "full"] and not is_expired:
                        return {
                            "success": True,
                            "can_reuse": False,
                            "reason": f"您已在有效 Team 中 ({team.team_name or team.id})，不可重复兑换",
                            "error": None
                        }

            # 6. 检查是否有过被封的记录
            has_banned_team = False
            for record in records:
                stmt = select(Team).where(Team.id == record.team_id)
                result = await db_session.execute(stmt)
                team = result.scalar_one_or_none()
                if team and team.status == "banned":
                    has_banned_team = True
                    break
            if has_banned_team:
                return {
                    "success": True,
                    "can_reuse": True,
                    "reason": "之前加入的 Team 已封号，可使用质保重复兑换",
                    "error": None
                }
            else:
                return {
                    "success": True,
                    "can_reuse": False,
                    "reason": "未找到被封号记录，且质保不支持正常过期或异常提示的重复兑换",
                    "error": None
                }

        except Exception as e:
            logger.error(f"验证质保码重复使用失败: {e}")
            return {
                "success": False,
                "can_reuse": False,
                "reason": None,
                "error": f"验证失败: {str(e)}"
            }


# 创建全局质保服务实例
warranty_service = WarrantyService()
