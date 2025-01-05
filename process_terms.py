import pandas as pd

# 1. [SMETA_용어등록 샘플(4)]의 데이터 추출
def extract_terms(input_file):
    df = pd.read_excel(input_file, sheet_name='용어목록', usecols=['No.', '용어명', '용어영문명'], engine='openpyxl')
    df.columns = ['No.', '용어명', '용어영문명']
    return df

# 2. D열 데이터를 '_'로 분리
def split_terms(term):
    return term.split('_')

# 3. [NIWP-DBA-DE_표준용어사전_V.0.92]에서 단어 정보 검색
def search_dictionary(dictionary_file, terms):
    dictionary_df = pd.read_excel(
        dictionary_file, 
        sheet_name='단어사전', 
        usecols=['표준 단어명', '영문 약어명', '단어 설명'], 
        engine='openpyxl'
    )
    dictionary_df.columns = ['표준 단어명', '영문 약어명', '단어 설명']  # 열 이름 재정의
    
    # 단어사전 데이터 출력
    print("단어사전 데이터프레임:")
    print(dictionary_df.head())  # 첫 5줄 출력

    results = {}
    for term in terms:
        match = dictionary_df[dictionary_df['영문 약어명'].str.lower() == term.lower()]
        if not match.empty:
            results[term] = {
                '표준 단어명': match.iloc[0]['표준 단어명'],
                '영문 약어명': match.iloc[0]['영문 약어명'],
                '단어 설명': match.iloc[0]['단어 설명']
            }
    return results

# 4. 결과를 새로운 엑셀 파일에 저장
def save_to_excel(output_file, extracted_data, matched_data):
    result = []
    for idx, row in extracted_data.iterrows():
        # '용어영문명'을 분리하여 각 단어 리스트 가져옴
        term_list = split_terms(row['용어영문명'])
        unique_terms = set(term_list)  # 중복된 단어 제거
        for term in unique_terms:
            if term in matched_data:  # 매칭된 단어만 처리
                matched_info = matched_data[term]
                result.append({
                    'No.': row['No.'],
                    '용어명': row['용어명'],
                    '용어영문명': row['용어영문명'],
                    '표준 단어명': matched_info['표준 단어명'],
                    '영문 약어명': matched_info['영문 약어명'],
                    '단어 설명': matched_info['단어 설명']
                })

    # 저장할 데이터가 없는 경우 메시지 출력
    if not result:
        print("저장할 데이터가 없습니다.")
        return

    # 데이터프레임 생성 및 엑셀 파일 저장
    result_df = pd.DataFrame(result)
    result_df.to_excel(output_file, index=False, engine='openpyxl')
    print(f"결과 파일이 {output_file}에 저장되었습니다.")

# 5. 실행
if __name__ == "__main__":
    input_file = "./project-folder/SMETA_용어등록 샘플(4).xlsx"
    dictionary_file = "./project-folder/NIWP-DBA-DE_표준용어사전_V.0.92.xlsx"
    output_file = "./project-folder/output.xlsx"

    # 데이터 추출
    extracted_data = extract_terms(input_file)
    print("추출된 데이터:", extracted_data.head()) # 첫 5줄 출력

    # 단어  분리
    terms = []
    for term in extracted_data['용어영문명']:
        terms.extend(split_terms(term))
    print("분리된 단어 리스트:", terms)

    # 단어사전에서 검색
    matched_data = search_dictionary(dictionary_file, terms)
    print("매칭된 데이터:", matched_data) #검색 결과 확인

    # 결과 저장
    save_to_excel(output_file, extracted_data, matched_data)
