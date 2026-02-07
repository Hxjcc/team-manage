"""
管理员路由
处理管理员面板的所有页面和操作
"""
import logging
from typing import Optional, List
from fastapi import APIRouter, Depends, HTTPException, status, Request
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
import json
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import func, and_, select
from pydantic import BaseModel, Field

from app.database import get_db
from app.models import Team, RedemptionCode
from app.dependencies.auth import require_admin
from app.services.team import TeamService
from app.services.redemption import RedemptionService
from app.utils.time_utils import get_now

logger = logging.getLogger(__name__)

# 创建路由器
router = APIRouter(
    prefix="/admin",
    tags=["admin"]
)

import json

# 服务实例
team_service = TeamService()
redemption_service = RedemptionService()


# 请求模型
class TeamImportRequest(BaseModel):
    """Team 导入请求"""
    import_type: str = Field(..., description="导入类型: single 或 batch")
    access_token: Optional[str] = Field(None, description="AT Token (单个导入)")
    refresh_token: Optional[str] = Field(None, description="Refresh Token (单个导入)")
    session_token: Optional[str] = Field(None, description="Session Token (单个导入)")
    client_id: Optional[str] = Field(None, description="Client ID (单个导入)")
    email: Optional[str] = Field(None, description="邮箱 (单个导入)")
    account_id: Optional[str] = Field(None, description="Account ID (单个导入)")
    content: Optional[str] = Field(None, description="批量导入内容")


class AddMemberRequest(BaseModel):
    """添加成员请求"""
    email: str = Field(..., description="成员邮箱")


class CodeGenerateRequest(BaseModel):
    """兑换码生成请求"""
    type: str = Field(..., description="生成类型: single 或 batch")
    code: Optional[str] = Field(None, description="自定义兑换码 (单个生成)")
    count: Optional[int] = Field(None, description="生成数量 (批量生成)")
    expires_days: Optional[int] = Field(None, description="有效期天数")
    team_id: Optional[int] = Field(None, description="绑定 Team ID (可选)")
    has_warranty: bool = Field(False, description="是否为质保兑换码")
    warranty_days: int = Field(30, description="质保天数")


class TeamUpdateRequest(BaseModel):
    """Team 更新请求"""
    email: Optional[str] = Field(None, description="新邮箱")
    account_id: Optional[str] = Field(None, description="新 Account ID")
    access_token: Optional[str] = Field(None, description="新 Access Token")
    refresh_token: Optional[str] = Field(None, description="新 Refresh Token")
    session_token: Optional[str] = Field(None, description="新 Session Token")
    client_id: Optional[str] = Field(None, description="新 Client ID")
    max_members: Optional[int] = Field(None, description="最大成员数")
    team_name: Optional[str] = Field(None, description="Team 名称")
    status: Optional[str] = Field(None, description="状态: active/full/expired/error/banned")


class BulkCodeDeleteRequest(BaseModel):
    """批量删除兑换码请求"""
    codes: List[str] = Field(..., description="兑换码列表")


@router.get("/", response_class=HTMLResponse)
async def admin_dashboard(
    request: Request,
    page: int = 1,
    search: Optional[str] = None,
    db: AsyncSession = Depends(get_db),
    current_user: dict = Depends(require_admin)
):
    """
    管理员面板首页
    """
    try:
        from app.main import templates
        logger.info(f"管理员访问控制台, search={search}, page={page}")

        # 设置每页数量
        per_page = 20
        
        # 获取 Team 列表 (分页)
        teams_result = await team_service.get_all_teams(db, page=page, per_page=per_page, search=search)
        
        # 使用高效的 COUNT 查询获取统计信息
        total_teams = (await db.execute(select(func.count(Team.id)))).scalar() or 0
        available_teams = (await db.execute(
            select(func.count(Team.id)).where(
                and_(Team.status == "active", Team.current_members < Team.max_members)
            )
        )).scalar() or 0
        total_codes = (await db.execute(select(func.count(RedemptionCode.id)))).scalar() or 0
        used_codes = (await db.execute(
            select(func.count(RedemptionCode.id)).where(RedemptionCode.status == "used")
        )).scalar() or 0

        stats = {
            "total_teams": total_teams,
            "available_teams": available_teams,
            "total_codes": total_codes,
            "used_codes": used_codes
        }

        return templates.TemplateResponse(
            "admin/index.html",
            {
                "request": request,
                "user": current_user,
                "active_page": "dashboard",
                "teams": teams_result.get("teams", []),
                "stats": stats,
                "search": search,
                "pagination": {
                    "current_page": teams_result.get("current_page", page),
                    "total_pages": teams_result.get("total_pages", 1),
                    "total": teams_result.get("total", 0),
                    "per_page": per_page
                }
            }
        )
    except Exception as e:
        logger.error(f"加载管理员面板失败: {e}")
        import traceback
        logger.error(traceback.format_exc())
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"加载管理员面板失败: {str(e)}"
        )


@router.post("/teams/{team_id}/delete")
async def delete_team(
    team_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: dict = Depends(require_admin)
):
    """
    删除 Team

    Args:
        team_id: Team ID
        db: 数据库会话
        current_user: 当前用户（需要登录）

    Returns:
        删除结果
    """
    try:
        logger.info(f"管理员删除 Team: {team_id}")

        result = await team_service.delete_team(team_id, db)

        if not result["success"]:
            return JSONResponse(
                status_code=status.HTTP_400_BAD_REQUEST,
                content=result
            )

        return JSONResponse(content=result)

    except Exception as e:
        logger.error(f"删除 Team 失败: {e}")
        return JSONResponse(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            content={
                "success": False,
                "error": f"删除 Team 失败: {str(e)}"
            }
        )


@router.get("/teams/{team_id}/info")
async def get_team_info(
    team_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: dict = Depends(require_admin)
):
    """获取 Team 详情 (包含解密后的 Token)"""
    try:
        result = await team_service.get_team_by_id(team_id, db)
        if not result["success"]:
            return JSONResponse(
                status_code=status.HTTP_404_NOT_FOUND,
                content=result
            )
        return JSONResponse(content=result)
    except Exception as e:
        return JSONResponse(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            content={"success": False, "error": str(e)}
        )


@router.post("/teams/{team_id}/update")
async def update_team(
    team_id: int,
    update_data: TeamUpdateRequest,
    db: AsyncSession = Depends(get_db),
    current_user: dict = Depends(require_admin)
):
    """更新 Team 信息"""
    try:
        result = await team_service.update_team(
            team_id=team_id,
            db_session=db,
            email=update_data.email,
            account_id=update_data.account_id,
            access_token=update_data.access_token,
            refresh_token=update_data.refresh_token,
            session_token=update_data.session_token,
            client_id=update_data.client_id,
            max_members=update_data.max_members,
            team_name=update_data.team_name,
            status=update_data.status
        )
        if not result["success"]:
            return JSONResponse(
                status_code=status.HTTP_400_BAD_REQUEST,
                content=result
            )
        return JSONResponse(content=result)
    except Exception as e:
        return JSONResponse(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            content={"success": False, "error": str(e)}
        )




@router.post("/teams/import")
async def team_import(
    import_data: TeamImportRequest,
    db: AsyncSession = Depends(get_db),
    current_user: dict = Depends(require_admin)
):
    """
    处理 Team 导入

    Args:
        import_data: 导入数据
        db: 数据库会话
        current_user: 当前用户（需要登录）

    Returns:
        导入结果
    """
    try:
        logger.info(f"管理员导入 Team: {import_data.import_type}")

        if import_data.import_type == "single":
            # 单个导入
            if not import_data.access_token:
                return JSONResponse(
                    status_code=status.HTTP_400_BAD_REQUEST,
                    content={
                        "success": False,
                        "error": "Access Token 不能为空"
                    }
                )

            result = await team_service.import_team_single(
                access_token=import_data.access_token,
                db_session=db,
                email=import_data.email,
                account_id=import_data.account_id,
                refresh_token=import_data.refresh_token,
                session_token=import_data.session_token,
                client_id=import_data.client_id
            )

            if not result["success"]:
                return JSONResponse(
                    status_code=status.HTTP_400_BAD_REQUEST,
                    content=result
                )

            return JSONResponse(content=result)

        elif import_data.import_type == "batch":
            # 批量导入使用 StreamingResponse
            async def progress_generator():
                async for status_item in team_service.import_team_batch(
                    text=import_data.content,
                    db_session=db
                ):
                    yield json.dumps(status_item, ensure_ascii=False) + "\n"

            return StreamingResponse(
                progress_generator(),
                media_type="application/x-ndjson"
            )

        else:
            return JSONResponse(
                status_code=status.HTTP_400_BAD_REQUEST,
                content={
                    "success": False,
                    "error": "无效的导入类型"
                }
            )

    except Exception as e:
        logger.error(f"导入 Team 失败: {e}")
        return JSONResponse(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            content={
                "success": False,
                "error": f"导入失败: {str(e)}"
            }
        )





@router.get("/teams/{team_id}/members/list")
async def team_members_list(
    team_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: dict = Depends(require_admin)
):
    """
    获取 Team 成员列表 (JSON)

    Args:
        team_id: Team ID
        db: 数据库会话
        current_user: 当前用户（需要登录）

    Returns:
        成员列表 JSON
    """
    try:
        # 获取成员列表
        result = await team_service.get_team_members(team_id, db)
        return JSONResponse(content=result)
    except Exception as e:
        logger.error(f"获取成员列表失败: {e}")
        return JSONResponse(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            content={
                "success": False,
                "error": f"获取成员列表失败: {str(e)}"
            }
        )


@router.post("/teams/{team_id}/members/add")
async def add_team_member(
    team_id: int,
    member_data: AddMemberRequest,
    db: AsyncSession = Depends(get_db),
    current_user: dict = Depends(require_admin)
):
    """
    添加 Team 成员

    Args:
        team_id: Team ID
        member_data: 成员数据
        db: 数据库会话
        current_user: 当前用户（需要登录）

    Returns:
        添加结果
    """
    try:
        logger.info(f"管理员添加成员到 Team {team_id}: {member_data.email}")

        result = await team_service.add_team_member(
            team_id=team_id,
            email=member_data.email,
            db_session=db
        )

        if not result["success"]:
            return JSONResponse(
                status_code=status.HTTP_400_BAD_REQUEST,
                content=result
            )

        return JSONResponse(content=result)

    except Exception as e:
        logger.error(f"添加成员失败: {e}")
        return JSONResponse(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            content={
                "success": False,
                "error": f"添加成员失败: {str(e)}"
            }
        )


@router.post("/teams/{team_id}/members/{user_id}/delete")
async def delete_team_member(
    team_id: int,
    user_id: str,
    db: AsyncSession = Depends(get_db),
    current_user: dict = Depends(require_admin)
):
    """
    删除 Team 成员

    Args:
        team_id: Team ID
        user_id: 用户 ID
        db: 数据库会话
        current_user: 当前用户（需要登录）

    Returns:
        删除结果
    """
    try:
        logger.info(f"管理员从 Team {team_id} 删除成员: {user_id}")

        result = await team_service.delete_team_member(
            team_id=team_id,
            user_id=user_id,
            db_session=db
        )

        if not result["success"]:
            return JSONResponse(
                status_code=status.HTTP_400_BAD_REQUEST,
                content=result
            )

        return JSONResponse(content=result)

    except Exception as e:
        logger.error(f"删除成员失败: {e}")
        return JSONResponse(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            content={
                "success": False,
                "error": f"删除成员失败: {str(e)}"
            }
        )


@router.post("/teams/{team_id}/invites/revoke")
async def revoke_team_invite(
    team_id: int,
    member_data: AddMemberRequest, # 使用相同的包含 email 的模型
    db: AsyncSession = Depends(get_db),
    current_user: dict = Depends(require_admin)
):
    """
    撤回 Team 邀请

    Args:
        team_id: Team ID
        member_data: 成员数据 (包含 email)
        db: 数据库会话
        current_user: 当前用户（需要登录）

    Returns:
        撤回结果
    """
    try:
        logger.info(f"管理员从 Team {team_id} 撤回邀请: {member_data.email}")

        result = await team_service.revoke_team_invite(
            team_id=team_id,
            email=member_data.email,
            db_session=db
        )

        if not result["success"]:
            return JSONResponse(
                status_code=status.HTTP_400_BAD_REQUEST,
                content=result
            )

        return JSONResponse(content=result)

    except Exception as e:
        logger.error(f"撤回邀请失败: {e}")
        return JSONResponse(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            content={
                "success": False,
                "error": f"撤回邀请失败: {str(e)}"
            }
        )


# ==================== Team 选项(用于生成兑换码绑定) ====================

@router.get("/teams/options")
async def get_team_options(
    db: AsyncSession = Depends(get_db),
    current_user: dict = Depends(require_admin)
):
    """
    获取可用 Team 列表(用于生成兑换码时绑定 Team)。

    Returns:
        { success: bool, teams: [...], error: str | null }
    """
    result = await team_service.get_available_teams_for_admin(db)
    if not result.get("success"):
        return JSONResponse(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            content={
                "success": False,
                "teams": [],
                "error": result.get("error") or "获取 Team 列表失败"
            }
        )

    return JSONResponse(content=result)


# ==================== 兑换码管理路由 ====================

@router.get("/codes", response_class=HTMLResponse)
async def codes_list_page(
    request: Request,
    page: int = 1,
    search: Optional[str] = None,
    db: AsyncSession = Depends(get_db),
    current_user: dict = Depends(require_admin)
):
    """
    兑换码列表页面

    Args:
        request: FastAPI Request 对象
        page: 页码
        search: 搜索关键词
        db: 数据库会话
        current_user: 当前用户（需要登录）

    Returns:
        兑换码列表页面 HTML
    """
    try:
        from app.main import templates

        logger.info(f"管理员访问兑换码列表页面, search={search}")

        # 获取兑换码 (分页)
        per_page = 50
        codes_result = await redemption_service.get_all_codes(db, page=page, per_page=per_page, search=search)
        codes = codes_result.get("codes", [])
        total_codes = codes_result.get("total", 0)
        total_pages = codes_result.get("total_pages", 1)
        current_page = codes_result.get("current_page", 1)

        # 为了统计数据，我们需要获取所有统计（或者增加统计接口）
        # 这里暂时获取全部用于统计
        all_codes_result = await redemption_service.get_all_codes(db, page=1, per_page=10000)
        all_codes = all_codes_result.get("codes", [])

        # 计算统计数据
        stats = {
            "total": total_codes,
            "unused": len([c for c in all_codes if c["status"] == "unused"]),
            # 已使用：普通码 used + 质保码 warranty_active
            "used": len([c for c in all_codes if c["status"] in ["used", "warranty_active"]]),
            "expired": len([c for c in all_codes if c["status"] == "expired"])
        }

        # 格式化日期时间
        from datetime import datetime
        for code in codes:
            if code.get("created_at"):
                dt = datetime.fromisoformat(code["created_at"])
                code["created_at"] = dt.strftime("%Y-%m-%d %H:%M")
            if code.get("expires_at"):
                dt = datetime.fromisoformat(code["expires_at"])
                code["expires_at"] = dt.strftime("%Y-%m-%d %H:%M")
            if code.get("used_at"):
                dt = datetime.fromisoformat(code["used_at"])
                code["used_at"] = dt.strftime("%Y-%m-%d %H:%M")

        # 绑定 Team 的价格信息 (用于按剩余时间展示)
        from sqlalchemy import select, func
        from app.models import Team, RedemptionRecord
        from app.utils.pricing import calculate_remaining_days, calculate_price_cents, format_price_yuan
        from app.utils.time_utils import get_now
        from datetime import timedelta

        team_ids = set()
        for c in codes:
            team_id = c.get("bound_team_id") or c.get("used_team_id")
            if team_id:
                team_ids.add(team_id)
        team_map = {}
        if team_ids:
            stmt = select(Team).where(Team.id.in_(list(team_ids)))
            result = await db.execute(stmt)
            teams = result.scalars().all()
            team_map = {t.id: t for t in teams}

        for code in codes:
            code["display_team_id"] = code.get("bound_team_id") or code.get("used_team_id")
            if code.get("bound_team_id"):
                code["display_team_source"] = "绑定"
            elif code.get("used_team_id"):
                code["display_team_source"] = "使用"
            else:
                code["display_team_source"] = None

            code["display_team_name"] = None
            code["display_remaining_days"] = None
            code["display_price_yuan"] = None

            team_id = code.get("display_team_id")
            if team_id:
                team = team_map.get(team_id)
                if team:
                    remaining_days = calculate_remaining_days(team.expires_at)
                    price_cents = calculate_price_cents(remaining_days)
                    code["display_team_name"] = team.team_name or f"Team {team.id}"
                    code["display_team_role"] = team.account_role
                    code["display_remaining_days"] = remaining_days
                    code["display_price_yuan"] = format_price_yuan(price_cents)

        # 质保剩余天数展示：默认跟随绑定 Team 到期（Team 到期质保即结束）
        # - 若该码已绑定/使用过 Team，则直接展示该 Team 的剩余天数
        # - 若没有 Team 信息，再回退到“首次使用时间 + 质保天数”
        warranty_codes = [
            c["code"]
            for c in codes
            if c.get("has_warranty") and c.get("status") != "unused" and not c.get("display_team_id")
        ]
        activation_map = {}
        if warranty_codes:
            stmt = (
                select(RedemptionRecord.code, func.min(RedemptionRecord.redeemed_at))
                .where(RedemptionRecord.code.in_(warranty_codes))
                .group_by(RedemptionRecord.code)
            )
            result = await db.execute(stmt)
            activation_map = {code: activated_at for code, activated_at in result.all()}

        now = get_now()
        for code in codes:
            code["warranty_remaining_days"] = None
            if not code.get("has_warranty"):
                continue

            # 有绑定/使用 Team：质保随 Team 到期
            if code.get("display_remaining_days") is not None:
                code["warranty_remaining_days"] = int(code["display_remaining_days"])
                continue

            if code.get("status") == "unused":
                continue

            activated_at = activation_map.get(code["code"])
            days = int(code.get("warranty_days") or 30)
            expiry_dt = None

            if activated_at:
                expiry_dt = activated_at + timedelta(days=days)
            elif code.get("warranty_expires_at"):
                try:
                    expiry_dt = datetime.fromisoformat(code["warranty_expires_at"])
                except Exception:
                    expiry_dt = None

            if expiry_dt:
                code["warranty_remaining_days"] = max((expiry_dt.date() - now.date()).days, 0)

        return templates.TemplateResponse(
            "admin/codes/index.html",
            {
                "request": request,
                "user": current_user,
                "active_page": "codes",
                "codes": codes,
                "stats": stats,
                "search": search,
                "pagination": {
                    "current_page": current_page,
                    "total_pages": total_pages,
                    "total": total_codes,
                    "per_page": per_page
                }
            }
        )

    except Exception as e:
        logger.error(f"加载兑换码列表页面失败: {e}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"加载页面失败: {str(e)}"
        )




@router.post("/codes/generate")
async def generate_codes(
    generate_data: CodeGenerateRequest,
    db: AsyncSession = Depends(get_db),
    current_user: dict = Depends(require_admin)
):
    """
    处理兑换码生成

    Args:
        generate_data: 生成数据
        db: 数据库会话
        current_user: 当前用户（需要登录）

    Returns:
        生成结果
    """
    try:
        logger.info(f"管理员生成兑换码: {generate_data.type}")

        if generate_data.type == "single":
            # 单个生成
            result = await redemption_service.generate_code_single(
                db_session=db,
                code=generate_data.code,
                bound_team_id=generate_data.team_id,
                expires_days=generate_data.expires_days,
                has_warranty=generate_data.has_warranty,
                warranty_days=generate_data.warranty_days
            )

            if not result["success"]:
                return JSONResponse(
                    status_code=status.HTTP_400_BAD_REQUEST,
                    content=result
                )

            return JSONResponse(content=result)

        elif generate_data.type == "batch":
            # 批量生成
            if not generate_data.count:
                return JSONResponse(
                    status_code=status.HTTP_400_BAD_REQUEST,
                    content={
                        "success": False,
                        "error": "生成数量不能为空"
                    }
                )

            result = await redemption_service.generate_code_batch(
                db_session=db,
                count=generate_data.count,
                bound_team_id=generate_data.team_id,
                expires_days=generate_data.expires_days,
                has_warranty=generate_data.has_warranty,
                warranty_days=generate_data.warranty_days
            )

            if not result["success"]:
                return JSONResponse(
                    status_code=status.HTTP_400_BAD_REQUEST,
                    content=result
                )

            return JSONResponse(content=result)

        else:
            return JSONResponse(
                status_code=status.HTTP_400_BAD_REQUEST,
                content={
                    "success": False,
                    "error": "无效的生成类型"
                }
            )

    except Exception as e:
        logger.error(f"生成兑换码失败: {e}")
        return JSONResponse(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            content={
                "success": False,
                "error": f"生成失败: {str(e)}"
            }
        )


@router.post("/codes/{code}/delete")
async def delete_code(
    code: str,
    db: AsyncSession = Depends(get_db),
    current_user: dict = Depends(require_admin)
):
    """
    删除兑换码

    Args:
        code: 兑换码
        db: 数据库会话
        current_user: 当前用户（需要登录）

    Returns:
        删除结果
    """
    try:
        logger.info(f"管理员删除兑换码: {code}")

        result = await redemption_service.delete_code(code, db)

        if not result["success"]:
            return JSONResponse(
                status_code=status.HTTP_400_BAD_REQUEST,
                content=result
            )

        return JSONResponse(content=result)

    except Exception as e:
        logger.error(f"删除兑换码失败: {e}")
        return JSONResponse(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            content={
                "success": False,
                "error": f"删除失败: {str(e)}"
            }
        )


@router.get("/codes/export")
async def export_codes(
    search: Optional[str] = None,
    db: AsyncSession = Depends(get_db),
    current_user: dict = Depends(require_admin)
):
    """
    导出兑换码为Excel文件

    Args:
        search: 搜索关键词
        db: 数据库会话
        current_user: 当前用户（需要登录）

    Returns:
        兑换码Excel文件
    """
    try:
        from fastapi.responses import Response
        from datetime import datetime
        import xlsxwriter
        from io import BytesIO

        logger.info("管理员导出兑换码为Excel")

        # 获取所有兑换码 (导出不分页，传入大数量)
        codes_result = await redemption_service.get_all_codes(db, page=1, per_page=100000, search=search)
        all_codes = codes_result.get("codes", [])
        
        # 结果可能带统计信息，我们只取 codes

        # 创建Excel文件到内存
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('兑换码列表')

        # 定义格式
        header_format = workbook.add_format({
            'bold': True,
            'fg_color': '#4F46E5',
            'font_color': 'white',
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })

        cell_format = workbook.add_format({
            'align': 'left',
            'valign': 'vcenter',
            'border': 1
        })

        # 设置列宽
        worksheet.set_column('A:A', 25)  # 兑换码
        worksheet.set_column('B:B', 12)  # 状态
        worksheet.set_column('C:C', 18)  # 创建时间
        worksheet.set_column('D:D', 18)  # 过期时间
        worksheet.set_column('E:E', 30)  # 使用者邮箱
        worksheet.set_column('F:F', 18)  # 使用时间
        worksheet.set_column('G:G', 12)  # 质保时长

        # 写入表头
        headers = ['兑换码', '状态', '创建时间', '过期时间', '使用者邮箱', '使用时间', '质保时长(天)']
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, header_format)

        # 写入数据
        for row, code in enumerate(all_codes, start=1):
            status_text = {
                'unused': '未使用',
                'used': '已使用',
                'expired': '已过期'
            }.get(code['status'], code['status'])

            worksheet.write(row, 0, code['code'], cell_format)
            worksheet.write(row, 1, status_text, cell_format)
            worksheet.write(row, 2, code.get('created_at', '-'), cell_format)
            worksheet.write(row, 3, code.get('expires_at', '永久有效'), cell_format)
            worksheet.write(row, 4, code.get('used_by_email', '-'), cell_format)
            worksheet.write(row, 5, code.get('used_at', '-'), cell_format)
            worksheet.write(row, 6, code.get('warranty_days', '-') if code.get('has_warranty') else '-', cell_format)

        # 关闭workbook
        workbook.close()

        # 获取Excel数据
        excel_data = output.getvalue()
        output.close()

        # 生成文件名
        filename = f"redemption_codes_{get_now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        # 返回Excel文件
        return Response(
            content=excel_data,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": f"attachment; filename={filename}"
            }
        )

    except Exception as e:
        logger.error(f"导出兑换码失败: {e}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"导出失败: {str(e)}"
        )


@router.post("/codes/bulk-delete")
async def bulk_delete_codes(
    delete_data: BulkCodeDeleteRequest,
    db: AsyncSession = Depends(get_db),
    current_user: dict = Depends(require_admin)
):
    """批量删除兑换码（仅未使用可删除）"""
    try:
        from sqlalchemy import select
        from app.models import RedemptionCode

        codes = [c.strip() for c in (delete_data.codes or []) if c and c.strip()]
        # 去重但保持顺序
        seen = set()
        codes = [c for c in codes if not (c in seen or seen.add(c))]

        if not codes:
            return JSONResponse(content={"success": True, "deleted": 0, "skipped": [], "not_found": []})

        stmt = select(RedemptionCode).where(RedemptionCode.code.in_(codes))
        result = await db.execute(stmt)
        objs = {obj.code: obj for obj in result.scalars().all()}

        deleted = []
        skipped = []
        not_found = []

        for code in codes:
            obj = objs.get(code)
            if not obj:
                not_found.append(code)
                continue
            if obj.status != "unused":
                skipped.append({"code": code, "reason": f"状态为 {obj.status}，不可删除"})
                continue

            await db.delete(obj)
            deleted.append(code)

        await db.commit()

        return JSONResponse(
            content={
                "success": True,
                "deleted": len(deleted),
                "deleted_codes": deleted,
                "skipped": skipped,
                "not_found": not_found,
                "error": None
            }
        )

    except Exception as e:
        logger.error(f"批量删除兑换码失败: {e}")
        await db.rollback()
        return JSONResponse(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            content={"success": False, "error": str(e)}
        )


@router.get("/records", response_class=HTMLResponse)
async def records_page(
    request: Request,
    email: Optional[str] = None,
    code: Optional[str] = None,
    team_id: Optional[str] = None,
    start_date: Optional[str] = None,
    end_date: Optional[str] = None,
    page: Optional[str] = "1",
    db: AsyncSession = Depends(get_db),
    current_user: dict = Depends(require_admin)
):
    """
    使用记录页面

    Args:
        request: FastAPI Request 对象
        email: 邮箱筛选
        code: 兑换码筛选
        team_id: Team ID 筛选
        start_date: 开始日期
        end_date: 结束日期
        page: 页码
        db: 数据库会话
        current_user: 当前用户（需要登录）

    Returns:
        使用记录页面 HTML
    """
    try:
        from app.main import templates
        from datetime import datetime, timedelta
        import math

        # 解析参数
        try:
            actual_team_id = int(team_id) if team_id and team_id.strip() else None
        except (ValueError, TypeError):
            actual_team_id = None
            
        try:
            page_int = int(page) if page and page.strip() else 1
        except (ValueError, TypeError):
            page_int = 1
            
        logger.info(f"管理员访问使用记录页面 (page={page_int})")

        # 获取记录 (支持邮箱、兑换码、Team ID 筛选)
        records_result = await redemption_service.get_all_records(
            db, 
            email=email, 
            code=code, 
            team_id=actual_team_id
        )
        all_records = records_result.get("records", [])

        # 仅由于日期范围筛选目前还在内存中处理，如果未来记录数极大可以移至数据库
        filtered_records = []
        for record in all_records:
            # 日期范围筛选
            if start_date or end_date:
                try:
                    record_date = datetime.fromisoformat(record["redeemed_at"]).date()

                    if start_date:
                        start = datetime.strptime(start_date, "%Y-%m-%d").date()
                        if record_date < start:
                            continue

                    if end_date:
                        end = datetime.strptime(end_date, "%Y-%m-%d").date()
                        if record_date > end:
                            continue
                except:
                    pass

            filtered_records.append(record)

        # 获取Team信息并关联到记录
        teams_result = await team_service.get_all_teams(db)
        teams = teams_result.get("teams", [])
        team_map = {team["id"]: team for team in teams}

        # 为记录添加Team名称
        for record in filtered_records:
            team = team_map.get(record["team_id"])
            record["team_name"] = team["team_name"] if team else None

        # 计算统计数据
        now = get_now()
        today_start = now.replace(hour=0, minute=0, second=0, microsecond=0)
        week_start = today_start - timedelta(days=today_start.weekday())
        month_start = today_start.replace(day=1)

        stats = {
            "total": len(filtered_records),
            "today": 0,
            "this_week": 0,
            "this_month": 0
        }

        for record in filtered_records:
            try:
                record_time = datetime.fromisoformat(record["redeemed_at"])
                if record_time >= today_start:
                    stats["today"] += 1
                if record_time >= week_start:
                    stats["this_week"] += 1
                if record_time >= month_start:
                    stats["this_month"] += 1
            except:
                pass

        # 分页
        per_page = 20
        total_records = len(filtered_records)
        total_pages = math.ceil(total_records / per_page) if total_records > 0 else 1

        # 确保页码有效
        if page_int < 1:
            page_int = 1
        if page_int > total_pages:
            page_int = total_pages

        start_idx = (page_int - 1) * per_page
        end_idx = start_idx + per_page
        paginated_records = filtered_records[start_idx:end_idx]

        # 格式化时间
        for record in paginated_records:
            try:
                dt = datetime.fromisoformat(record["redeemed_at"])
                record["redeemed_at"] = dt.strftime("%Y-%m-%d %H:%M:%S")
            except:
                pass

        return templates.TemplateResponse(
            "admin/records/index.html",
            {
                "request": request,
                "user": current_user,
                "active_page": "records",
                "records": paginated_records,
                "stats": stats,
                "filters": {
                    "email": email,
                    "code": code,
                    "team_id": team_id,
                    "start_date": start_date,
                    "end_date": end_date
                },
                "pagination": {
                    "current_page": page_int,
                    "total_pages": total_pages,
                    "total": total_records,
                    "per_page": per_page
                }
            }
        )

    except Exception as e:
        logger.error(f"获取使用记录失败: {e}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"获取使用记录失败: {str(e)}"
        )


@router.post("/records/{record_id}/withdraw")
async def withdraw_record(
    record_id: int,
    db: AsyncSession = Depends(get_db),
    current_user: dict = Depends(require_admin)
):
    """
    撤中使用记录 (管理员功能)

    Args:
        record_id: 记录 ID
        db: 数据库会话
        current_user: 当前用户（需要登录）

    Returns:
        结果 JSON
    """
    try:
        logger.info(f"管理员请求撤回记录: {record_id}")
        result = await redemption_service.withdraw_record(record_id, db)

        if not result["success"]:
            return JSONResponse(
                status_code=status.HTTP_400_BAD_REQUEST,
                content=result
            )

        return JSONResponse(content=result)

    except Exception as e:
        logger.error(f"撤回记录失败: {e}")
        return JSONResponse(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            content={
                "success": False,
                "error": f"撤回失败: {str(e)}"
            }
        )


@router.get("/settings", response_class=HTMLResponse)
async def settings_page(
    request: Request,
    db: AsyncSession = Depends(get_db),
    current_user: dict = Depends(require_admin)
):
    """
    系统设置页面

    Args:
        request: FastAPI Request 对象
        db: 数据库会话
        current_user: 当前用户（需要登录）

    Returns:
        系统设置页面 HTML
    """
    try:
        from app.main import templates
        from app.services.settings import settings_service

        logger.info("管理员访问系统设置页面")

        # 获取当前配置
        flaresolverr_config = await settings_service.get_flaresolverr_config(db)
        log_level = await settings_service.get_log_level(db)

        return templates.TemplateResponse(
            "admin/settings/index.html",
            {
                "request": request,
                "user": current_user,
                "active_page": "settings",
                "flaresolverr_enabled": flaresolverr_config["enabled"],
                "flaresolverr_url": flaresolverr_config["url"],
                "log_level": log_level
            }
        )

    except Exception as e:
        logger.error(f"获取系统设置失败: {e}")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"获取系统设置失败: {str(e)}"
        )


class LogLevelRequest(BaseModel):
    """日志级别请求"""
    level: str = Field(..., description="日志级别")


class FlareSolverrConfigRequest(BaseModel):
    """FlareSolverr 配置请求"""
    enabled: bool = Field(..., description="是否启用 FlareSolverr")
    url: str = Field("", description="FlareSolverr 服务地址")


@router.post("/settings/log-level")
async def update_log_level(
    log_data: LogLevelRequest,
    db: AsyncSession = Depends(get_db),
    current_user: dict = Depends(require_admin)
):
    """
    更新日志级别

    Args:
        log_data: 日志级别数据
        db: 数据库会话
        current_user: 当前用户（需要登录）

    Returns:
        更新结果
    """
    try:
        from app.services.settings import settings_service

        logger.info(f"管理员更新日志级别: {log_data.level}")

        # 更新日志级别
        success = await settings_service.update_log_level(db, log_data.level)

        if success:
            return JSONResponse(content={"success": True, "message": "日志级别已保存"})
        else:
            return JSONResponse(
                status_code=status.HTTP_400_BAD_REQUEST,
                content={"success": False, "error": "无效的日志级别"}
            )

    except Exception as e:
        logger.error(f"更新日志级别失败: {e}")
        return JSONResponse(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            content={"success": False, "error": f"更新失败: {str(e)}"}
        )


@router.post("/settings/flaresolverr")
async def update_flaresolverr_config(
    config_data: FlareSolverrConfigRequest,
    db: AsyncSession = Depends(get_db),
    current_user: dict = Depends(require_admin)
):
    """
    更新 FlareSolverr 配置

    Args:
        config_data: FlareSolverr 配置数据
        db: 数据库会话
        current_user: 当前用户（需要登录）

    Returns:
        更新结果
    """
    try:
        from app.services.settings import settings_service

        logger.info(f"管理员更新 FlareSolverr 配置: enabled={config_data.enabled}, url={config_data.url}")

        # 验证 URL 格式
        if config_data.enabled and config_data.url:
            url = config_data.url.strip()
            if not (url.startswith("http://") or url.startswith("https://")):
                return JSONResponse(
                    status_code=status.HTTP_400_BAD_REQUEST,
                    content={
                        "success": False,
                        "error": "FlareSolverr 地址格式错误,应为 http://host:port"
                    }
                )

        # 更新配置
        success = await settings_service.update_flaresolverr_config(
            db,
            config_data.enabled,
            config_data.url.strip() if config_data.url else ""
        )

        if success:
            # 清理 ChatGPT 服务的会话和 CF cookies 缓存
            from app.services.chatgpt import chatgpt_service
            await chatgpt_service.clear_session()

            return JSONResponse(content={"success": True, "message": "FlareSolverr 配置已保存"})
        else:
            return JSONResponse(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                content={"success": False, "error": "保存失败"}
            )

    except Exception as e:
        logger.error(f"更新 FlareSolverr 配置失败: {e}")
        return JSONResponse(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            content={"success": False, "error": f"更新失败: {str(e)}"}
        )


