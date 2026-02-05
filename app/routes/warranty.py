"""
质保相关路由
处理用户质保查询请求
"""
from fastapi import APIRouter, Depends, HTTPException
from pydantic import BaseModel, EmailStr, field_validator
from typing import Optional
from sqlalchemy.ext.asyncio import AsyncSession

from app.database import get_db
from app.services.warranty import warranty_service
from app.utils.code_utils import normalize_code_input

router = APIRouter(
    prefix="/warranty",
    tags=["warranty"]
)


class WarrantyCheckRequest(BaseModel):
    """质保查询请求"""
    email: Optional[EmailStr] = None
    code: Optional[str] = None

    @field_validator("code", mode="before")
    @classmethod
    def _normalize_code(cls, value: str | None) -> str | None:
        return normalize_code_input(value)


class WarrantyCheckRecord(BaseModel):
    """质保查询单条记录"""
    code: str
    has_warranty: bool
    warranty_valid: bool
    warranty_expires_at: Optional[str]
    status: str
    used_at: Optional[str]
    team_id: Optional[int]
    team_name: Optional[str]
    team_status: Optional[str]
    team_expires_at: Optional[str]
    email: Optional[str] = None


class WarrantyCheckResponse(BaseModel):
    """质保查询响应"""
    success: bool
    has_warranty: bool
    warranty_valid: bool
    warranty_expires_at: Optional[str]
    banned_teams: list
    can_reuse: bool
    original_code: Optional[str]
    records: list[WarrantyCheckRecord] = []
    message: Optional[str]
    error: Optional[str]


@router.post("/check", response_model=WarrantyCheckResponse)
async def check_warranty(
    request: WarrantyCheckRequest,
    db_session: AsyncSession = Depends(get_db)
):
    """
    检查质保状态
    
    用户可以通过邮箱或兑换码查询质保状态
    """
    try:
        # 验证至少提供一个参数
        if not request.email and not request.code:
            raise HTTPException(
                status_code=400,
                detail="必须提供邮箱或兑换码"
            )
        
        # 调用质保服务
        result = await warranty_service.check_warranty_status(
            db_session,
            email=request.email,
            code=request.code
        )
        
        if not result["success"]:
            raise HTTPException(
                status_code=500,
                detail=result.get("error", "查询失败")
            )
        
        return WarrantyCheckResponse(
            success=True,
            has_warranty=result.get("has_warranty", False),
            warranty_valid=result.get("warranty_valid", False),
            warranty_expires_at=result.get("warranty_expires_at"),
            banned_teams=result.get("banned_teams", []),
            can_reuse=result.get("can_reuse", False),
            original_code=result.get("original_code"),
            records=result.get("records", []),
            message=result.get("message"),
            error=None
        )
        
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"查询质保状态失败: {str(e)}"
        )
