"""
response_grouper.py v2
주관식 응답 공통응답 처리 — 핵심 원칙:
  - 2건 이상 반복되거나 유사한 응답만 '공통응답'으로 묶음
  - 단독 응답(1건)은 그룹화하지 않고 개별 표시
  - 공통응답은 가나다 오름차순 정렬
  - 개별복수의 "기타" 한 덩어리로 묶지 않음
"""
import re


# ─── 규칙 기반 유사 의미 사전 ───
MIN_GROUP_COUNT = 2  # 공통응답 최소 건수

SYNONYM_GROUPS = [
    {
        "label": "교육이 유익하고 만족스러웠습니다.",
        "keywords": ["유익", "유용", "도움", "만족", "좋았", "좋습니다", "좋은", "훌륭", "최고",
                     "감사", "유의미", "가치", "의미있", "의미 있"],
    },
    {
        "label": "강사의 역량과 전달력이 우수했습니다.",
        "keywords": ["강사", "강의", "선생", "설명", "전달", "쉽게", "명확"],
    },
    {
        "label": "교육 내용과 구성이 적절했습니다.",
        "keywords": ["내용", "구성", "커리큘럼", "프로그램", "주제", "수준", "적절", "알차"],
    },
    {
        "label": "실무에 바로 적용할 수 있는 교육이었습니다.",
        "keywords": ["실무", "현업", "적용", "실제", "실질", "바로", "즉시"],
    },
    {
        "label": "교육 시간이나 분량 조정이 필요합니다.",
        "keywords": ["시간", "분량", "짧", "길", "부족", "더 많", "늘려"],
    },
    {
        "label": "지속적인 교육 운영을 희망합니다.",
        "keywords": ["다음", "또 하", "재교육", "지속", "계속", "반복", "정기"],
    },
    {
        "label": "교육 환경 및 시설이 쾌적했습니다.",
        "keywords": ["환경", "시설", "장소", "공간", "쾌적", "편리"],
    },
    {
        "label": "교육 내용의 심화 또는 개선이 필요합니다.",
        "keywords": ["개선", "보완", "심화", "발전", "향상", "업그레이드"],
    },
    {
        "label": "동료와의 교류 및 네트워킹이 유익했습니다.",
        "keywords": ["동료", "네트워크", "교류", "소통", "팀워크", "협력"],
    },
]


def group_responses_rule_based(answers):
    """
    규칙 기반 공통응답 그룹핑
    - 2건 이상 같은 카테고리에 해당하는 응답만 '공통응답'으로 묶음
    - 나머지는 개별 원문 그대로 표시
    """
    groups = {}       # label → [answers]
    individuals = []  # 어느 그룹에도 안 속하는 개별 응답
    
    for answer in answers:
        matched = False
        for sg in SYNONYM_GROUPS:
            if any(kw in answer for kw in sg["keywords"]):
                lbl = sg["label"]
                if lbl not in groups:
                    groups[lbl] = []
                groups[lbl].append(answer)
                matched = True
                break
        if not matched:
            individuals.append(answer)
    
    result = []
    
    # ── 공통응답: 2건 이상인 그룹만 ──
    for lbl, ans_list in groups.items():
        if len(ans_list) >= MIN_GROUP_COUNT:
            result.append({
                "label": lbl,
                "answers": ans_list,
                "count": len(ans_list),
                "is_common": True,
            })
        else:
            # 1건짜리는 개별로 내림
            individuals.extend(ans_list)
    
    # 가나다 오름차순
    result.sort(key=lambda x: x["label"])
    
    # 공통응답N 라벨
    for i, item in enumerate(result):
        item["common_id"] = f"공통응답{i+1}"
    
    # ── 개별 응답: 원문 그대로 각각 표시 ──
    for ans in individuals:
        result.append({
            "label": ans,          # 원문이 라벨
            "answers": [ans],
            "count": 1,
            "is_common": False,
            "common_id": "",       # 공통응답 번호 없음
        })
    
    return result


def generate_grouping_prompt(open_ended_item):
    """Gemini AI에 보낼 공통응답 그룹핑 프롬프트 (v2)"""
    answers = open_ended_item.get("answers", [])
    if not answers:
        return None
    
    answers_text = "\n".join([f"{i+1}. {a}" for i, a in enumerate(answers[:80])])
    
    prompt = f"""다음은 교육 만족도 설문의 주관식 응답입니다.

문항: {open_ended_item.get('label', '')}
응답 수: {len(answers)}건

응답 목록:
{answers_text}

[중요 규칙]
1. **2건 이상** 비슷한 의미를 가진 응답들만 묶어 '공통응답'으로 처리하세요.
2. 1건뿐인 개별 응답은 그룹으로 묶지 말고 원문 그대로 individual로 분류하세요.
3. **[절대 금지]**: 단순 강사 이름, 인명, 부서명, 또는 카테고리 레이블(과정명, 장소 등)을 '공통응답' 테마로 묶지 마세요. 이들은 각각 개별 응답으로 유지해야 합니다.
4. 묶인 공통응답에는 대표 문장을 만드세요 (자연스러운 완전한 문장).
5. 공통응답 그룹은 가나다 오름차순으로 정렬하세요.
6. 응답 원문은 절대 수정하지 마세요.
7. "기타 의견"이라는 포괄 그룹을 만들지 마세요 — 각각 개별로 표기하세요.

출력 형식 (JSON):
{{
    "groups": [
        {{
            "label": "공통응답 대표 문장 (2건 이상일 때만)",
            "count": 해당_응답_수,
            "answers": ["원문1", "원문2"],
            "is_common": true
        }}
    ],
    "individuals": [
        "단독 응답 원문 1",
        "단독 응답 원문 2"
    ]
}}

JSON만 출력하세요."""
    return prompt


async def group_responses_ai(open_ended_item, ai_engine):
    """AI 기반 공통응답 그룹핑 (2건 이상만 묶음)"""
    answers = open_ended_item.get("answers", [])
    
    if not ai_engine or not ai_engine.enabled:
        return group_responses_rule_based(answers)
    
    prompt = generate_grouping_prompt(open_ended_item)
    if not prompt:
        return []
    
    try:
        result = await ai_engine.generate_narrative(prompt)
        
        if isinstance(result, dict) and "groups" in result:
            groups = result.get("groups", [])
            individuals = result.get("individuals", [])
            
            # 공통응답: 2건 이상만
            valid_groups = [g for g in groups if g.get("count", 0) >= MIN_GROUP_COUNT]
            valid_groups.sort(key=lambda x: x.get("label", ""))
            for i, g in enumerate(valid_groups):
                g["common_id"] = f"공통응답{i+1}"
                g["is_common"] = True
            
            # 개별 응답: 원문 각각
            individual_items = []
            for ans in individuals:
                individual_items.append({
                    "label": ans,
                    "answers": [ans],
                    "count": 1,
                    "is_common": False,
                    "common_id": "",
                })
            
            return valid_groups + individual_items
    except Exception as e:
        print(f"AI 그룹핑 실패: {e}")
    
    return group_responses_rule_based(answers)


async def process_all_open_ended(open_ended_list, ai_engine=None):
    """모든 주관식 문항에 공통응답 그룹핑 적용 — 3중 검증 포함"""
    result = []
    
    # ── 사전 준비: 전체 문항별 응답 집합 (교차 오염 탐지용) ──
    answer_ownership = {}  # answer_text → oe_id
    for oe in open_ended_list:
        for ans in oe.get("answers", []):
            ans_key = ans.strip().lower()
            if ans_key not in answer_ownership:
                answer_ownership[ans_key] = oe.get("id", "")
    
    for oe in open_ended_list:
        oe_id = oe.get("id", "")
        oe_label = oe.get("label", "")
        answers = oe.get("answers", [])
        
        if not answers:
            result.append({**oe, "groups": [], "individual_count": 0, "common_count": 0})
            continue
        
        # ── 검증 1: 응답이 이 문항 소속인지 확인 ──
        valid_answers = []
        for ans in answers:
            ans_key = ans.strip().lower()
            owner = answer_ownership.get(ans_key, oe_id)
            if owner == oe_id:
                valid_answers.append(ans)
            else:
                print(f"[QA-CHECK1] '{ans[:30]}...' → {oe_id}에서 제거 (원래 {owner} 소속)")
        
        if not valid_answers:
            result.append({**oe, "groups": [], "individual_count": 0, "common_count": 0})
            continue
        
        # 필터링된 응답으로 oe 교체
        oe_clean = {**oe, "answers": valid_answers}
        
        # 문항별 독립 그룹핑
        groups = await group_responses_ai(oe_clean, ai_engine)
        
        # ── 검증 2: 그룹 내 응답이 원본 answers에 실재하는지 확인 ──
        answer_set = set(a.strip().lower() for a in valid_answers)
        for g in groups:
            g_answers = g.get("answers", [])
            verified = [a for a in g_answers if a.strip().lower() in answer_set]
            if len(verified) != len(g_answers):
                removed = len(g_answers) - len(verified)
                print(f"[QA-CHECK2] {oe_id} 그룹 '{g.get('label','')}': {removed}건 잘못된 응답 제거")
            g["answers"] = verified
            g["count"] = len(verified)
        
        # 빈 그룹 제거
        groups = [g for g in groups if g["count"] > 0]
        
        # ── 검증 3: 전체 그룹 응답 수 vs 원본 응답 수 일관성 ──
        total_grouped = sum(g["count"] for g in groups)
        if total_grouped > len(valid_answers) * 1.5:
            print(f"[QA-CHECK3] ⚠ {oe_id} 그룹 합계({total_grouped}) > 원본({len(valid_answers)})*1.5 — 비정상")
        
        common_groups = [g for g in groups if g.get("is_common")]
        individual_groups = [g for g in groups if not g.get("is_common")]
        
        result.append({
            **oe,
            "answers": valid_answers,  # 검증된 응답만
            "groups": groups,
            "common_groups": common_groups,
            "individual_responses": individual_groups,
            "common_count": len(common_groups),
            "individual_count": len(individual_groups),
            "total_grouped": sum(g["count"] for g in common_groups),
        })
    
    return result
