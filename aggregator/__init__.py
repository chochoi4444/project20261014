"""기준가격대장 총자산 취합기: 운용사별 보고서에서 총자산을 뽑아 당사 양식에 채운다."""
from .extract import ReportResult, extract_report
from .fill import fill_form

__all__ = ["ReportResult", "extract_report", "fill_form"]
