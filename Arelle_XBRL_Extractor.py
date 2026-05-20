"""
XBRL Fact 및 Presentation 정보 통합 추출기

XBRL 파일(.xbrl / .xml)을 열어 Fact 값과 Presentation 계층 구조를
함께 추출한 뒤, Excel 파일로 저장하는 스크립트.

출력 컬럼:
    Label_Korean, Label_English, Seq, ContextRef, Decimals, Value,
    Pres_Level, Pres_Role_Name, Pres_Role_URI, Pres_Parent_Label_KR,
    Pres_Parent_Name, Concept_Name, Concept_ID, Abstract,
    Substitution_Group, Type, Period_Type, Balance
"""

def main():
    """메인 함수 — 파일 선택 → 파싱 → 매칭 → Excel 저장 순으로 실행"""

    # ------------------------------------------------------------------ #
    # 1. 파일 선택 
    # ------------------------------------------------------------------ #
    from tkinter import Tk
    from tkinter.filedialog import askopenfilename

    Tk().withdraw()  # tkinter 루트 창 숨김 (파일 선택창만 표시)
    print("\nXBRL 파일을 선택하세요...")
    xbrl_file_path = askopenfilename(
        title="XBRL 파일 선택",
        filetypes=[("XBRL files", "*.xbrl"), ("XML files", "*.xml"), ("All files", "*.*")]
    )

    if not xbrl_file_path:
        print("파일이 선택되지 않았습니다.")
        return

    # ------------------------------------------------------------------ #
    # 2. 나머지 라이브러리 로드
    #    파일 선택 이후에 로드하여 초기 실행 속도를 높임
    # ------------------------------------------------------------------ #
    print("\n필요한 라이브러리 로딩 중...")
    from arelle import Cntlr          # XBRL 파싱 엔진
    import pandas as pd               # 데이터 프레임 / Excel 출력
    from pathlib import Path          # 경로 처리
    from datetime import datetime     # 타임스탬프 생성
    from openpyxl.styles import Font  # Excel 셀 스타일
    from collections import defaultdict

    print("로딩 완료\n")

    # ================================================================== #
    # [내부 함수 1] get_concept_labels_fast
    #   concept 객체에서 한국어·영어 레이블을 빠르게 추출
    #   레이블이 없으면 qname.localName 으로 대체
    # ================================================================== #
    def get_concept_labels_fast(concept):
        ko_label = en_label = None
        try:
            ko_label = concept.label(lang='ko')
            en_label = concept.label(lang='en')
        except:
            pass
        # 한국어 레이블이 없을 경우 concept의 로컬명을 사용
        if not ko_label:
            ko_label = concept.qname.localName if concept.qname else None
        return ko_label, en_label

    # ================================================================== #
    # [내부 함수 2] extract_role_korean_names
    #   model_xbrl.roleTypes 에서 role URI → 한국어 명칭 딕셔너리를 생성
    #   definition 문자열에 '|' 구분자가 있으면 첫 번째 토큰만 사용
    #   예) "연결재무상태표 | Consolidated Balance Sheet"
    #        → "연결재무상태표"
    # ================================================================== #
    def extract_role_korean_names(model_xbrl):
        role_names = {}
        if not hasattr(model_xbrl, 'roleTypes') or not model_xbrl.roleTypes:
            return role_names
        for role_uri, role_type_list in model_xbrl.roleTypes.items():
            try:
                # role_type_list 가 리스트인 경우 첫 번째 요소 사용
                role_type = role_type_list[0] if isinstance(role_type_list, list) and role_type_list else role_type_list
                if role_type and hasattr(role_type, 'definition') and role_type.definition:
                    definition = role_type.definition
                    role_names[role_uri] = definition.split('|')[0].strip() if '|' in definition else definition.strip()
            except:
                continue
        return role_names

    # ================================================================== #
    # [내부 함수 3] build_presentation_map_optimized
    #   Presentation Linkbase(parent-child 관계)를 순회하여
    #   concept별 표시 정보(role, level, parent)를 수집
    #
    #   반환값:
    #     presentation_map  : {concept_key: [pres_info, ...]}
    #     all_concepts_info : {concept_key: {레이블, 타입 등 메타데이터}}
    # ================================================================== #
    def build_presentation_map_optimized(model_xbrl, role_names):
        presentation_map = defaultdict(list)  # concept → pres 정보 리스트
        all_concepts_info = {}                # concept → 메타데이터

        if not hasattr(model_xbrl, 'relationshipSet') or not hasattr(model_xbrl, 'roleTypes'):
            return presentation_map, all_concepts_info

        for role_uri in model_xbrl.roleTypes.keys():
            try:
                # 해당 role의 parent-child 관계 집합 조회
                rel_set = model_xbrl.relationshipSet(
                    "http://www.xbrl.org/2003/arcrole/parent-child",
                    role_uri
                )

                if not rel_set or not rel_set.modelRelationships:
                    continue  # 관계가 없는 role은 스킵

                role_korean_name = role_names.get(role_uri)
                children_map = defaultdict(list)  # parent_key → [child_key, ...]
                all_concepts = set()              # (concept_key, concept 객체) 쌍

                # --- 관계 순회: parent-child 맵 구성 --- #
                for rel in rel_set.modelRelationships:
                    parent = rel.fromModelObject
                    child = rel.toModelObject
                    parent_key = str(parent.qname)
                    child_key = str(child.qname)
                    children_map[parent_key].append(child_key)
                    all_concepts.add((parent_key, parent))
                    all_concepts.add((child_key, child))

                # --- BFS로 각 concept의 계층 레벨(depth) 계산 --- #
                level_map = {}
                # 부모가 없는 concept = 루트
                roots = {
                    key for key, concept in all_concepts
                    if key not in [child for children in children_map.values() for child in children]
                }
                queue = [(root, 0) for root in roots]
                processed = set()

                while queue:
                    concept_key, level = queue.pop(0)
                    if concept_key in processed:
                        continue
                    processed.add(concept_key)
                    # 같은 concept이 여러 경로로 참조될 경우 최대 depth 사용
                    level_map[concept_key] = max(level_map.get(concept_key, 0), level)
                    for child_key in children_map.get(concept_key, []):
                        queue.append((child_key, level + 1))

                # --- 자식 → 부모 역방향 맵 구성 (첫 번째 부모만 저장) --- #
                parent_map = {}
                for parent_key, child_keys in children_map.items():
                    for child_key in child_keys:
                        if child_key not in parent_map:
                            parent_map[child_key] = parent_key

                # --- 각 concept의 pres_info 저장 및 메타데이터 수집 --- #
                for concept_key, concept in all_concepts:
                    level = level_map.get(concept_key, 0)
                    parent_key = parent_map.get(concept_key)

                    pres_info = {
                        'Pres_Role_URI': role_uri,
                        'Pres_Role_Name': role_korean_name,
                        'Pres_Level': level,
                        'Pres_Parent_Name': parent_key
                    }

                    presentation_map[concept_key].append(pres_info)

                    # 메타데이터는 concept 당 최초 1회만 수집
                    if concept_key not in all_concepts_info:
                        ko_label, en_label = get_concept_labels_fast(concept)
                        all_concepts_info[concept_key] = {
                            'Label_Korean': ko_label,
                            'Label_English': en_label,
                            'Concept_Name': concept_key,
                            'Concept_ID': getattr(concept, 'id', None),
                            'Abstract': getattr(concept, 'isAbstract', False),
                            'Substitution_Group': str(concept.substitutionGroup) if hasattr(concept, 'substitutionGroup') and concept.substitutionGroup else None,
                            'Type': str(concept.type.qname) if hasattr(concept, 'type') and concept.type and hasattr(concept.type, 'qname') else None,
                            'Period_Type': getattr(concept, 'periodType', None),
                            'Balance': getattr(concept, 'balance', None)
                        }
            except:
                continue  # 개별 role 파싱 오류는 무시하고 계속 진행

        # --- pres_info에 부모 concept의 한국어 레이블 후처리 추가 --- #
        for concept_key, pres_list in presentation_map.items():
            for pres_info in pres_list:
                parent_key = pres_info.get('Pres_Parent_Name')
                if parent_key and parent_key in all_concepts_info:
                    pres_info['Pres_Parent_Label_KR'] = all_concepts_info[parent_key]['Label_Korean']
                else:
                    pres_info['Pres_Parent_Label_KR'] = None

        return dict(presentation_map), all_concepts_info

    # ------------------------------------------------------------------ #
    # 3. XBRL 파일 로드 (Arelle 엔진)
    #    validate=False : 유효성 검증 생략 → 속도 향상
    #    logger level 50 : CRITICAL 이상만 출력 → 콘솔 노이즈 억제
    # ------------------------------------------------------------------ #
    print("파일 로딩 중...")
    start_time = datetime.now()

    ctrl = Cntlr.Cntlr(logFileName="logToPrint")
    ctrl.logger.setLevel(50)  # CRITICAL 레벨만 출력 (INFO/WARNING 숨김)
    model_xbrl = ctrl.modelManager.load(xbrl_file_path, validate=False)

    if model_xbrl is None:
        print("데이터 추출 실패")
        return

    print(f"Fact {len(model_xbrl.facts)}개 로드 완료")

    # ------------------------------------------------------------------ #
    # 4. Role 정보 및 Presentation 구조 추출
    # ------------------------------------------------------------------ #
    print("Role 정보 추출 중...")
    role_names = extract_role_korean_names(model_xbrl)

    print("Presentation 구조 분석 중...")
    presentation_map, all_concepts_info = build_presentation_map_optimized(model_xbrl, role_names)
    print(f"Concept {len(all_concepts_info)}개 발견")

    # ------------------------------------------------------------------ #
    # 5. Fact 값을 concept 단위로 그룹화
    #    {concept_key: [fact, fact, ...]} 형태로 인덱싱하여
    #    이후 매칭 루프에서 O(1) 조회 가능하게 함
    # ------------------------------------------------------------------ #
    print("Fact 정보 매칭 중...")
    fact_by_concept = defaultdict(list)
    for fact in model_xbrl.facts:
        try:
            concept_name = str(fact.concept.qname) if fact.concept.qname else None
            if concept_name:
                fact_by_concept[concept_name].append(fact)
        except:
            continue

    # ------------------------------------------------------------------ #
    # 6. Fact × Presentation 크로스 조인 → 행 데이터 생성
    #
    #    Fact가 있는 concept : fact 수 × pres 수 만큼 행 생성
    #    Fact가 없는 concept : pres 수 만큼 행 생성 (Value=None)
    #    → Presentation 구조에 있지만 값이 없는 항목도 누락 없이 포함
    # ------------------------------------------------------------------ #
    fact_data_list = []
    seq = 1  # 전체 행에 걸쳐 순번 부여 (1부터 시작)

    for concept_name, concept_info in all_concepts_info.items():
        pres_list = presentation_map.get(concept_name, [])
        facts = fact_by_concept.get(concept_name, [])

        if facts:
            # Fact가 존재하는 경우: (fact, pres_info) 조합마다 행 추가
            for fact in facts:
                for pres_info in pres_list:
                    fact_data_list.append({
                        'Label_Korean': concept_info['Label_Korean'],
                        'Label_English': concept_info['Label_English'],
                        'Seq': seq,
                        'ContextRef': getattr(fact, 'contextID', None),   # 기간/시점 컨텍스트 ID
                        'Decimals': getattr(fact, 'decimals', None),      # 소수점 자리수
                        'Value': getattr(fact, 'value', None),            # 실제 값
                        'Pres_Level': pres_info.get('Pres_Level'),
                        'Pres_Role_Name': pres_info.get('Pres_Role_Name'),
                        'Pres_Role_URI': pres_info.get('Pres_Role_URI'),
                        'Pres_Parent_Label_KR': pres_info.get('Pres_Parent_Label_KR'),
                        'Pres_Parent_Name': pres_info.get('Pres_Parent_Name'),
                        'Concept_Name': concept_info['Concept_Name'],
                        'Concept_ID': concept_info['Concept_ID'],
                        'Abstract': concept_info['Abstract'],
                        'Substitution_Group': concept_info['Substitution_Group'],
                        'Type': concept_info['Type'],
                        'Period_Type': concept_info['Period_Type'],
                        'Balance': concept_info['Balance']
                    })
                    seq += 1
        else:
            # Fact가 없는 경우: 구조 정보만 기록 (Value 관련 컬럼은 None)
            for pres_info in pres_list:
                fact_data_list.append({
                    'Label_Korean': concept_info['Label_Korean'],
                    'Label_English': concept_info['Label_English'],
                    'Seq': seq,
                    'ContextRef': None,
                    'Decimals': None,
                    'Value': None,
                    'Pres_Level': pres_info.get('Pres_Level'),
                    'Pres_Role_Name': pres_info.get('Pres_Role_Name'),
                    'Pres_Role_URI': pres_info.get('Pres_Role_URI'),
                    'Pres_Parent_Label_KR': pres_info.get('Pres_Parent_Label_KR'),
                    'Pres_Parent_Name': pres_info.get('Pres_Parent_Name'),
                    'Concept_Name': concept_info['Concept_Name'],
                    'Concept_ID': concept_info['Concept_ID'],
                    'Abstract': concept_info['Abstract'],
                    'Substitution_Group': concept_info['Substitution_Group'],
                    'Type': concept_info['Type'],
                    'Period_Type': concept_info['Period_Type'],
                    'Balance': concept_info['Balance']
                })
                seq += 1

    fact_df = pd.DataFrame(fact_data_list)

    if fact_df.empty:
        print("데이터 추출 실패")
        return

    processing_time = (datetime.now() - start_time).total_seconds()

    # ------------------------------------------------------------------ #
    # 7. Excel 파일 저장
    #    저장 경로: ~/Downloads/XBRL_Complete_{원본파일명}_{타임스탬프}.xlsx
    # ------------------------------------------------------------------ #
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = Path(xbrl_file_path).stem  # 확장자 제외 파일명
    output_path = Path.home() / "Downloads" / f"XBRL_Complete_{file_name}_{timestamp}.xlsx"

    print("\n" + "=" * 80)
    print(f"Excel 파일 저장 중: {output_path}")

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        fact_df.to_excel(writer, index=False, sheet_name='Complete_Data')
        worksheet = writer.sheets['Complete_Data']

        # 시트 기본 줌 70%로 설정 (데이터 컬럼이 많아 가독성 확보)
        worksheet.sheet_view.zoomScale = 70

        # 헤더 행 폰트 크기 10pt 적용
        font_10 = Font(size=10)
        for cell in worksheet[1]:
            cell.font = font_10

        # 각 컬럼 너비를 헤더 + 상위 100개 셀 기준으로 자동 조정 (최대 50)
        for column in worksheet.columns:
            column_letter = column[0].column_letter
            header_length = len(str(column[0].value))
            sample_lengths = [len(str(cell.value)) for cell in list(column)[1:min(101, len(column))]]
            max_length = max([header_length] + sample_lengths) if sample_lengths else header_length
            adjusted_width = min(max_length + 2, 50)  # 최소 여백 +2, 최대 50
            worksheet.column_dimensions[column_letter].width = adjusted_width

    print(f"✓ 저장 완료")
    print("\n" + "=" * 80)
    print("프로그램을 종료하려면 Enter를 누르세요...")
    input()


if __name__ == "__main__":
    # 제목을 가장 먼저 출력 (import 없이)
    print("=" * 80)
    print("XBRL Fact & Presentation 통합 추출기")
    print("=" * 80)

    main()