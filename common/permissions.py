"""
共用權限檢查函數
所有 App 共用的角色判斷邏輯集中於此
"""
from common.constants import GROUP_NAMES


def is_simple_applicant(user):
    """檢查是否為簡易申請人員（包含主管與管理員）"""
    return user.groups.filter(
        name__in=[GROUP_NAMES['APPLICANT'], GROUP_NAMES['APPLICANT_SUPERVISOR']]
    ).exists() or user.is_superuser


def is_simple_dispatcher(user):
    """檢查是否為簡易撥料人員（包含主管與管理員）"""
    return user.groups.filter(
        name__in=[GROUP_NAMES['DISPATCHER'], GROUP_NAMES['DISPATCHER_SUPERVISOR']]
    ).exists() or user.is_superuser


def is_applicant_supervisor(user):
    """檢查是否為申請人員主管"""
    return user.groups.filter(name=GROUP_NAMES['APPLICANT_SUPERVISOR']).exists()


def is_dispatcher_supervisor(user):
    """檢查是否為撥料人員主管"""
    return user.groups.filter(name=GROUP_NAMES['DISPATCHER_SUPERVISOR']).exists()


def is_supervisor(user):
    """檢查是否為任一主管角色"""
    return user.groups.filter(
        name__in=[GROUP_NAMES['APPLICANT_SUPERVISOR'], GROUP_NAMES['DISPATCHER_SUPERVISOR']]
    ).exists()
