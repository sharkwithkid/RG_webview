# core/errors.py

from dataclasses import dataclass
from typing import Dict


@dataclass(frozen=True)
class ErrorDef:
    """
    RGclassmove 파이프라인에서 사용하는 에러 코드 정의.

    - code   : 에러 코드(대문자 + 언더스코어)
    - domain : 큰 영역(COMMON/SCAN/DB/ROSTER/FRESHMEN/TRANSFER/WITHDRAW/REGISTER/NOTICE/RESOURCES/TEACHER 등)
    - summary: 한국어 요약 (로그/문서용 요약 설명)
    """
    code: str
    domain: str
    summary: str


ERROR_DEFS: Dict[str, ErrorDef] = {

    # =========================
    # COMMON
    # =========================
    "UNEXPECTED_ERROR": {
    "domain": "COMMON",
    "dev_summary": "예상하지 못한 내부 예외가 발생했습니다.",
    "user_message": "처리 중 알 수 없는 오류가 발생했습니다. 다시 시도해 주세요. 문제가 계속되면 관리자에게 문의해 주세요."
    },

    "SCAN_NOT_OK": {
        "domain": "COMMON",
        "dev_summary": "scan 단계에서 ok=False 상태로 실행이 시도되었습니다.",
        "user_message": "파일 스캔 단계에서 문제가 발생했습니다. 먼저 스캔 결과를 확인해 주세요."
    },

    "TEMPLATE_IO_ERROR": {
        "domain": "COMMON",
        "dev_summary": "템플릿 파일 로딩 또는 저장 중 I/O 오류 발생.",
        "user_message": "양식 파일을 읽거나 저장하는 중 오류가 발생했습니다. 파일이 열려 있는지 확인해 주세요."
    },

    # =========================
    # WORK ROOT / FOLDER
    # =========================

    "WORK_ROOT_NOT_FOUND": {
    "domain": "WORK_ROOT",
    "dev_summary": "지정된 작업 루트 경로가 존재하지 않습니다.",
    "user_message": "작업 폴더를 찾을 수 없습니다. 경로를 다시 확인해 주세요."
    },

    "SCHOOL_FOLDER_NOT_FOUND": {
        "domain": "WORK_ROOT",
        "dev_summary": "학교명을 포함한 하위 폴더를 찾지 못했습니다.",
        "user_message": "작업 폴더 안에서 해당 학교 폴더를 찾지 못했습니다. 폴더명을 확인해 주세요."
    },

    "MULTIPLE_SCHOOL_FOLDER_MATCH": {
        "domain": "WORK_ROOT",
        "dev_summary": "학교명과 일치하는 폴더가 2개 이상 발견되었습니다.",
        "user_message": "학교명과 일치하는 폴더가 여러 개 발견되었습니다. 폴더명을 정리해 주세요."
    },

    # =========================
    # DB
    # =========================

    "DB_FILE_NOT_FOUND": {
    "domain": "DB",
    "dev_summary": "DB(.xlsb) 파일을 찾지 못했습니다.",
    "user_message": "DB 파일을 찾을 수 없습니다. resources/DB 폴더를 확인해 주세요."
    },

    "DB_SCHOOL_NOT_FOUND": {
    "domain": "DB",
    "dev_summary": "DB에서 해당 학교명을 찾지 못했습니다.",
    "user_message": "DB 파일에서 해당 학교 정보를 찾지 못했습니다. 학교명이 정확한지 확인해 주세요."
    },

    "DB_DOMAIN_MISSING": {
    "domain": "DB",
    "dev_summary": "DB에 학교 도메인 정보가 존재하지 않습니다.",
    "user_message": "학교 도메인 정보가 등록되어 있지 않습니다. 관리자에게 문의해 주세요."
    },

    # =========================
    # FRESHMEN
    # =========================
    "FRESHMEN_FILE_NOT_FOUND": {
        "domain": "FRESHMEN",
        "dev_summary": "신입생 파일을 찾지 못했습니다.",
        "user_message": "신입생 파일을 찾지 못했습니다. 파일명을 확인해 주세요."
    },

    "FRESHMEN_HEADER_NOT_FOUND": {
        "domain": "FRESHMEN",
        "dev_summary": "신입생 파일에서 헤더 행을 감지하지 못했습니다.",
        "user_message": "신입생 파일에서 첫 행의 열 제목을 인식하지 못했습니다. 파일 형식을 확인해 주세요."
    },

    "FRESHMEN_HEADER_MISSING_REQUIRED": {
        "domain": "FRESHMEN",
        "dev_summary": "신입생 파일 헤더에서 필수 열(학년/반/이름)이 누락되었습니다.",
        "user_message": "신입생 파일 첫 행에 '학년', '반', '이름' 열이 있는지 확인해 주세요."
    },

    "FRESHMEN_INVALID_GRADE_VALUE": {
        "domain": "FRESHMEN",
        "dev_summary": "신입생 파일에서 학년 값을 숫자로 변환하지 못했습니다.",
        "user_message": "신입생 파일의 학년 값이 올바르지 않습니다. 숫자 형태로 입력해 주세요."
    },

    # =========================
    # TRANSFER
    # =========================
    "TRANSFER_FILE_NOT_FOUND": {
        "domain": "TRANSFER",
        "dev_summary": "전입생 파일을 찾지 못했습니다.",
        "user_message": "전입생 파일을 찾지 못했습니다. 파일이 있는지 확인해 주세요."
    },

    "TRANSFER_HEADER_MISSING_REQUIRED": {
        "domain": "TRANSFER",
        "dev_summary": "전입생 파일 헤더에서 필수 열이 누락되었습니다.",
        "user_message": "전입생 파일의 열 제목이 올바른지 확인해 주세요."
    },

    "TRANSFER_STUDENT_NOT_IN_ROSTER": {
        "domain": "TRANSFER",
        "dev_summary": "전입생이 학생명부에 존재하지 않습니다.",
        "user_message": "전입생 명단 중 학생명부에서 찾을 수 없는 학생이 있습니다. 명부를 확인해 주세요."
    },

    # =========================
    # WITHDRAW
    # =========================

    "WITHDRAW_FILE_NOT_FOUND": {
    "domain": "WITHDRAW",
    "dev_summary": "전출생 파일을 찾지 못했습니다.",
    "user_message": "전출생 파일을 찾지 못했습니다."
    },

    "WITHDRAW_HEADER_MISSING_REQUIRED": {
        "domain": "WITHDRAW",
        "dev_summary": "전출생 파일 헤더에서 필수 열이 누락되었습니다.",
        "user_message": "전출생 파일의 열 제목이 올바른지 확인해 주세요."
    },

    "WITHDRAW_STUDENT_NOT_IN_ROSTER": {
        "domain": "WITHDRAW",
        "dev_summary": "전출생이 학생명부에 존재하지 않습니다.",
        "user_message": "전출생 명단 중 학생명부에서 찾을 수 없는 학생이 있습니다."
    },

    # =========================
    # ROSTER
    # =========================
    "ROSTER_FILE_NOT_FOUND": {
        "domain": "ROSTER",
        "dev_summary": "학생명부 파일을 찾지 못했습니다.",
        "user_message": "학생명부 파일을 찾지 못했습니다."
    },

    "ROSTER_HEADER_NOT_FOUND": {
        "domain": "ROSTER",
        "dev_summary": "학생명부 헤더를 감지하지 못했습니다.",
        "user_message": "학생명부 파일의 열 제목을 인식하지 못했습니다."
    },

    "ROSTER_BASIS_DATE_MISMATCH": {
        "domain": "ROSTER",
        "dev_summary": "명부 기준일이 작업일과 일치하지 않습니다.",
        "user_message": "학생명부 기준일이 현재 작업일과 다릅니다. 최신 명부인지 확인해 주세요."
    },

    # =========================
    # REGISTER TEMPLATE
    # =========================
    "REGISTER_TEMPLATE_NOT_FOUND": {
    "domain": "REGISTER_TEMPLATE",
    "dev_summary": "등록 양식 파일을 찾지 못했습니다.",
    "user_message": "등록 작업용 양식 파일을 찾지 못했습니다."
    },

    "REGISTER_TEMPLATE_STRUCTURE_INVALID": {
        "domain": "REGISTER_TEMPLATE",
        "dev_summary": "등록 양식 파일 구조가 예상과 다릅니다.",
        "user_message": "등록 양식 파일 형식이 올바르지 않습니다. 기본 양식을 사용해 주세요."
    },


    # =========================
    # NOTICE TEMPLATE
    # =========================

    "NOTICE_TEMPLATE_NOT_FOUND": {
    "domain": "NOTICE_TEMPLATE",
    "dev_summary": "안내문 텍스트 파일을 찾지 못했습니다.",
    "user_message": "안내문 양식 파일을 찾지 못했습니다."
    },

    "NOTICE_TEMPLATE_RENDER_ERROR": {
        "domain": "NOTICE_TEMPLATE",
        "dev_summary": "안내문 치환 중 오류가 발생했습니다.",
        "user_message": "안내문 생성 중 오류가 발생했습니다."
    },