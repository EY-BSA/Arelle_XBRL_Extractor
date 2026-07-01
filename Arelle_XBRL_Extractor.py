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

from tkinter import Tk
from tkinter.filedialog import askopenfilename
from arelle import Cntlr
from pathlib import Path
from datetime import datetime
from collections import defaultdict, deque
import pandas as pd
from openpyxl.styles import Font


def get_concept_labels(concept):
    """concept의 한국어·영어 레이블 반환. 한국어 레이블 없으면 localName으로 대체."""
    ko_label = en_label = None
    try:
        ko_label = concept.label(lang='ko')
        en_label = concept.label(lang='en')
    except:
        pass
    if not ko_label:
        ko_label = concept.qname.localName if concept.qname else None
    return ko_label, en_label


def extract_role_korean_names(model_xbrl):
    """role URI → 한국어 명칭 딕셔너리 반환.
    definition이 '연결재무상태표 | Consolidated...' 형태면 '|' 앞 토큰만 사용."""
    role_names = {}
    if not hasattr(model_xbrl, 'roleTypes') or not model_xbrl.roleTypes:
        return role_names
    for role_uri, role_type_list in model_xbrl.roleTypes.items():
        try:
            role_type = role_type_list[0] if isinstance(role_type_list, list) and role_type_list else role_type_list
            if role_type and hasattr(role_type, 'definition') and role_type.definition:
                definition = role_type.definition
                role_names[role_uri] = definition.split('|')[0].strip() if '|' in definition else definition.strip()
        except:
            continue
    return role_names


def build_presentation_map(model_xbrl, role_names):
    """Presentation Linkbase를 순회해 concept별 표시 정보와 메타데이터를 수집.

    반환값:
        presentation_map  : {concept_key: [pres_info, ...]}  — role별 계층 정보
        all_concepts_info : {concept_key: 레이블·타입 등 메타데이터}
    """
    presentation_map = defaultdict(list)
    all_concepts_info = {}

    if not hasattr(model_xbrl, 'relationshipSet') or not hasattr(model_xbrl, 'roleTypes'):
        return presentation_map, all_concepts_info

    for role_uri in model_xbrl.roleTypes.keys():
        try:
            rel_set = model_xbrl.relationshipSet(
                "http://www.xbrl.org/2003/arcrole/parent-child",
                role_uri
            )

            if not rel_set or not rel_set.modelRelationships:
                continue

            role_korean_name = role_names.get(role_uri)
            children_map = defaultdict(list)  # parent_key → [child_key, ...]
            all_concepts = set()              # (concept_key, concept 객체) 쌍

            for rel in rel_set.modelRelationships:
                parent = rel.fromModelObject
                child = rel.toModelObject
                parent_key = str(parent.qname)
                child_key = str(child.qname)
                children_map[parent_key].append(child_key)
                all_concepts.add((parent_key, parent))
                all_concepts.add((child_key, child))

            # 부모가 없는 concept = 루트. set으로 만들어 O(1) 멤버십 체크
            child_set = {child for children in children_map.values() for child in children}
            roots = {key for key, _ in all_concepts if key not in child_set}
            # BFS로 각 concept의 계층 레벨(depth) 계산
            queue = deque((root, 0) for root in roots)
            level_map = {}
            processed = set()

            while queue:
                concept_key, level = queue.popleft()
                if concept_key in processed:
                    continue
                processed.add(concept_key)
                # 같은 concept이 여러 경로로 참조될 경우 최대 depth 사용
                level_map[concept_key] = max(level_map.get(concept_key, 0), level)
                for child_key in children_map.get(concept_key, []):
                    queue.append((child_key, level + 1))

            # 자식 → 부모 역방향 맵 (첫 번째 부모만 저장)
            parent_map = {}
            for parent_key, child_keys in children_map.items():
                for child_key in child_keys:
                    if child_key not in parent_map:
                        parent_map[child_key] = parent_key

            for concept_key, concept in all_concepts:
                presentation_map[concept_key].append({
                    'Pres_Role_URI': role_uri,
                    'Pres_Role_Name': role_korean_name,
                    'Pres_Level': level_map.get(concept_key, 0),
                    'Pres_Parent_Name': parent_map.get(concept_key)
                })

                # 메타데이터는 concept당 최초 1회만 수집
                if concept_key not in all_concepts_info:
                    ko_label, en_label = get_concept_labels(concept)
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
            continue

    # pres_info에 부모 concept의 한국어 레이블 후처리 (build 완료 후 일괄 처리)
    for concept_key, pres_list in presentation_map.items():
        for pres_info in pres_list:
            parent_key = pres_info.get('Pres_Parent_Name')
            pres_info['Pres_Parent_Label_KR'] = (
                all_concepts_info[parent_key]['Label_Korean']
                if parent_key and parent_key in all_concepts_info else None
            )

    return dict(presentation_map), all_concepts_info


def main():
    Tk().withdraw()
    print("\nXBRL 파일을 선택하세요...")
    xbrl_file_path = askopenfilename(
        title="XBRL 파일 선택",
        filetypes=[("XBRL files", "*.xbrl"), ("XML files", "*.xml"), ("All files", "*.*")]
    )

    if not xbrl_file_path:
        print("파일이 선택되지 않았습니다.")
        return

    print("파일 로딩 중...")

    ctrl = Cntlr.Cntlr(logFileName="logToPrint")
    ctrl.logger.setLevel(50)  # CRITICAL 이상만 출력 — INFO/WARNING 콘솔 노이즈 억제
    model_xbrl = ctrl.modelManager.load(xbrl_file_path, validate=False)  # 유효성 검증 생략으로 속도 향상

    if model_xbrl is None:
        print("데이터 추출 실패")
        return

    print(f"Fact {len(model_xbrl.facts)}개 로드 완료")

    print("Role 정보 추출 중...")
    role_names = extract_role_korean_names(model_xbrl)

    print("Presentation 구조 분석 중...")
    presentation_map, all_concepts_info = build_presentation_map(model_xbrl, role_names)
    print(f"Concept {len(all_concepts_info)}개 발견")

    # Fact를 concept 단위로 그룹화 → 이후 매칭 루프에서 O(1) 조회
    print("Fact 정보 매칭 중...")
    fact_by_concept = defaultdict(list)
    for fact in model_xbrl.facts:
        try:
            concept_name = str(fact.concept.qname) if fact.concept.qname else None
            if concept_name:
                fact_by_concept[concept_name].append(fact)
        except:
            continue

    fact_data_list = []
    seq = 1

    for concept_name, concept_info in all_concepts_info.items():
        pres_list = presentation_map.get(concept_name, [])
        # Fact가 없는 concept는 [None]으로 대체해 구조 정보만 포함한 행을 생성
        for fact in (fact_by_concept.get(concept_name) or [None]):
            for pres_info in pres_list:
                fact_data_list.append({
                    'Label_Korean': concept_info['Label_Korean'],
                    'Label_English': concept_info['Label_English'],
                    'Seq': seq,
                    'ContextRef': getattr(fact, 'contextID', None),
                    'Decimals': getattr(fact, 'decimals', None),
                    'Value': getattr(fact, 'value', None),
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

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_name = Path(xbrl_file_path).stem
    output_path = Path.home() / "Downloads" / f"XBRL_Complete_{file_name}_{timestamp}.xlsx"

    print("\n" + "=" * 80)
    print(f"Excel 파일 저장 중: {output_path}")

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        fact_df.to_excel(writer, index=False, sheet_name='Complete_Data')
        worksheet = writer.sheets['Complete_Data']
        worksheet.sheet_view.zoomScale = 70  # 컬럼이 많아 기본 줌을 70%로 축소

        for cell in worksheet[1]:
            cell.font = Font(size=10)

        # 헤더 + 상위 100개 셀 기준으로 컬럼 너비 자동 조정 (최대 50)
        for column in worksheet.columns:
            column_letter = column[0].column_letter
            header_length = len(str(column[0].value))
            sample_lengths = [len(str(cell.value)) for cell in list(column)[1:min(101, len(column))]]
            max_length = max([header_length] + sample_lengths) if sample_lengths else header_length
            worksheet.column_dimensions[column_letter].width = min(max_length + 2, 50)

    print(f"✓ 저장 완료")
    print("\n" + "=" * 80)
    print("프로그램을 종료하려면 Enter를 누르세요...")
    input()


if __name__ == "__main__":
    print("=" * 80)
    print("XBRL Fact & Presentation 통합 추출기")
    print("=" * 80)

    main()
