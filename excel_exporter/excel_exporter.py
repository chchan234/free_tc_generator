"""
엑셀 내보내기 모듈: 테스트케이스 및 검증 결과를 엑셀 파일로 내보내기
"""

import os
import pandas as pd
from typing import List, Dict, Any
from datetime import datetime

def export_to_excel(testcases: List[Dict[str, Any]], output_dir: str = "data/output") -> str:
    """
    테스트케이스를 엑셀 파일로 내보내기
    
    Args:
        testcases: 내보낼 테스트케이스 리스트
        output_dir: 엑셀 파일을 저장할 디렉토리
        
    Returns:
        생성된 엑셀 파일 경로
    """
    # 출력 디렉토리 생성
    os.makedirs(output_dir, exist_ok=True)
    
    # 타임스탬프 생성
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # 출력 파일 경로
    output_file = os.path.join(output_dir, f"testcases_{timestamp}.xlsx")
    
    # DataFrame 생성
    df = pd.DataFrame(testcases)
    
    # 필수 열이 없는 경우 빈 열 추가
    required_columns = ["대분류", "중분류", "소분류", "확인내용", "플랫폼", "비고"]
    for col in required_columns:
        if col not in df.columns:
            df[col] = ""
    
    # 열 순서 정렬
    df = df[required_columns]
    
    # 엑셀 파일로 내보내기
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="테스트케이스", index=False)
        
        # 엑셀 서식 지정
        workbook = writer.book
        worksheet = writer.sheets["테스트케이스"]
        
        # 열 너비 설정
        worksheet.column_dimensions['A'].width = 15  # 대분류
        worksheet.column_dimensions['B'].width = 15  # 중분류
        worksheet.column_dimensions['C'].width = 15  # 소분류
        worksheet.column_dimensions['D'].width = 40  # 확인내용
        worksheet.column_dimensions['E'].width = 15  # 플랫폼
        worksheet.column_dimensions['F'].width = 20  # 비고
    
    return output_file

def export_validation_results(testcases: List[Dict[str, Any]], validation_results: List[Dict[str, Any]], output_dir: str = "data/output") -> str:
    """
    테스트케이스와 검증 결과를 함께 엑셀 파일로 내보내기
    
    Args:
        testcases: 내보낼 테스트케이스 리스트
        validation_results: 테스트케이스 검증 결과 리스트
        output_dir: 엑셀 파일을 저장할 디렉토리
        
    Returns:
        생성된 엑셀 파일 경로
    """
    # 출력 디렉토리 생성
    os.makedirs(output_dir, exist_ok=True)
    
    # 타임스탬프 생성
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # 출력 파일 경로
    output_file = os.path.join(output_dir, f"testcases_validated_{timestamp}.xlsx")
    
    # 테스트케이스 DataFrame 생성
    tc_df = pd.DataFrame(testcases)
    
    # 필수 열이 없는 경우 빈 열 추가
    required_columns = ["대분류", "중분류", "소분류", "확인내용", "플랫폼", "비고"]
    for col in required_columns:
        if col not in tc_df.columns:
            tc_df[col] = ""
    
    # 열 순서 정렬
    tc_df = tc_df[required_columns]
    
    # 검증 결과 DataFrame 생성
    val_df = pd.DataFrame([
        {
            "정확성": result["정확성"],
            "완전성": result["완전성"],
            "명확성": result["명확성"],
            "플랫폼_적합성": result["플랫폼_적합성"],
            "총점": result["총점"],
            "통과_여부": result["통과_여부"],
            "개선_제안": result["개선_제안"]
        }
        for result in validation_results
    ])
    
    # 엑셀 파일로 내보내기
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        tc_df.to_excel(writer, sheet_name="테스트케이스", index=False)
        val_df.to_excel(writer, sheet_name="검증 결과", index=False)
        
        # 엑셀 서식 지정
        workbook = writer.book
        
        # 테스트케이스 시트 서식
        worksheet = writer.sheets["테스트케이스"]
        worksheet.column_dimensions['A'].width = 15  # 대분류
        worksheet.column_dimensions['B'].width = 15  # 중분류
        worksheet.column_dimensions['C'].width = 15  # 소분류
        worksheet.column_dimensions['D'].width = 40  # 확인내용
        worksheet.column_dimensions['E'].width = 15  # 플랫폼
        worksheet.column_dimensions['F'].width = 20  # 비고
        
        # 검증 결과 시트 서식
        worksheet = writer.sheets["검증 결과"]
        worksheet.column_dimensions['A'].width = 10  # 정확성
        worksheet.column_dimensions['B'].width = 10  # 완전성
        worksheet.column_dimensions['C'].width = 10  # 명확성
        worksheet.column_dimensions['D'].width = 15  # 플랫폼_적합성
        worksheet.column_dimensions['E'].width = 10  # 총점
        worksheet.column_dimensions['F'].width = 10  # 통과_여부
        worksheet.column_dimensions['G'].width = 40  # 개선_제안
    
    return output_file