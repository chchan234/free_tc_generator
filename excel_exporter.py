"""
엑셀 내보내기 모듈: 생성된 테스트케이스를 엑셀 파일로 내보내기
"""

import os
from typing import List, Dict, Any
import pandas as pd
from datetime import datetime

def export_to_excel(testcases: List[Dict[str, str]], output_file: str = None) -> str:
    """
    테스트케이스를 엑셀 파일로 내보내기
    
    Args:
        testcases: 내보낼 테스트케이스 목록
        output_file: 출력 파일 경로 (None인 경우 자동 생성)
        
    Returns:
        생성된 엑셀 파일 경로
    """
    # 기본 열 정의
    columns = ["대분류", "중분류", "소분류", "확인내용", "플랫폼", "비고"]
    
    # DataFrame 생성
    df = pd.DataFrame(testcases)
    
    # 열 순서 지정
    if all(col in df.columns for col in columns):
        df = df[columns]
    
    # 출력 파일명 자동 생성
    if output_file is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = os.path.join(os.getcwd(), "data", "output")
        os.makedirs(output_dir, exist_ok=True)
        output_file = os.path.join(output_dir, f"testcases_{timestamp}.xlsx")
    
    # 엑셀 파일로 저장
    df.to_excel(output_file, index=False, sheet_name="테스트케이스")
    
    # 엑셀 파일 서식 조정
    with pd.ExcelWriter(output_file, engine='openpyxl', mode='a') as writer:
        # 기존 시트 제거
        writer.book.remove(writer.book.active)
        
        # 새로 작성
        df.to_excel(writer, index=False, sheet_name="테스트케이스")
        
        # 워크시트 가져오기
        worksheet = writer.sheets["테스트케이스"]
        
        # 열 너비 조정
        for i, column in enumerate(df.columns):
            column_width = max(len(column) * 1.5, df[column].astype(str).map(len).max() * 1.2)
            worksheet.column_dimensions[chr(65 + i)].width = column_width
    
    return output_file

def export_validation_results(validation_results: List[Dict[str, Any]], output_file: str = None) -> str:
    """
    검증 결과를 엑셀 파일로 내보내기
    
    Args:
        validation_results: 내보낼 검증 결과 목록
        output_file: 출력 파일 경로 (None인 경우 자동 생성)
        
    Returns:
        생성된 엑셀 파일 경로
    """
    # 결과 데이터 처리
    processed_results = []
    
    for result in validation_results:
        testcase = result.pop("testcase", {})
        # 테스트케이스와 검증 결과 통합
        processed_result = {**testcase, **result}
        processed_results.append(processed_result)
    
    # 기본 열 정의
    columns = ["대분류", "중분류", "소분류", "확인내용", "플랫폼", "비고", 
               "정확성", "완전성", "명확성", "플랫폼_적합성", "총점", "개선_제안", "통과_여부"]
    
    # DataFrame 생성
    df = pd.DataFrame(processed_results)
    
    # 열 순서 지정
    columns = [col for col in columns if col in df.columns]
    df = df[columns]
    
    # 출력 파일명 자동 생성
    if output_file is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = os.path.join(os.getcwd(), "data", "output")
        os.makedirs(output_dir, exist_ok=True)
        output_file = os.path.join(output_dir, f"validated_testcases_{timestamp}.xlsx")
    
    # 엑셀 파일로 저장
    df.to_excel(output_file, index=False, sheet_name="검증결과")
    
    # 엑셀 파일 서식 조정
    with pd.ExcelWriter(output_file, engine='openpyxl', mode='a') as writer:
        # 기존 시트 제거
        writer.book.remove(writer.book.active)
        
        # 새로 작성
        df.to_excel(writer, index=False, sheet_name="검증결과")
        
        # 워크시트 가져오기
        worksheet = writer.sheets["검증결과"]
        
        # 열 너비 조정
        for i, column in enumerate(df.columns):
            column_width = max(len(column) * 1.5, df[column].astype(str).map(len).max() * 1.2)
            worksheet.column_dimensions[chr(65 + i)].width = column_width
    
    return output_file