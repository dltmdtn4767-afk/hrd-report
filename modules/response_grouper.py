"""
response_grouper.py
주관식 응답을 AI(Gemini) 또는 규칙 기반으로 유사 의미 묶음 처리
- 그룹화 후 가나다 오름차순 정렬
- 공통응답N 형식으로 라벨
"""
import re


# ─── 규칙 기반 유사 의미 사전 ───
SYNONYM_GROUPS = [
    # 긍정/만족
    {
        "label": "교육이 유익하고 만족스러웠습니다.",
        "keywords": ["유익", "유용", "도움", "만족", "좋았", "좋습니다", "좋은", "훌륭", "최고",
                     "감사", "유의미", "가치", "의미있", "의미 있"],
    },
    # 강사 칭찬
    {
        "label": "강사의 역량과 전달력이 우수했습니다.",
        "keywords": ["강사", "강의", "선생", "설명", "전달", "쉽게", "명확"],
    },
    # 내용/구성
    {
        "label": "교육 내용과 구성이 적절했습니다.",
        "keywords": ["내용", "구성", "커리큘럼", "프로그램", "주제", "수준", "적절", "알차"],
    },
    # 실무 적용
    {
        "label": "실무에 바로 적용할 수 있는 교육이었습니다.",
        "keywords": ["실무", "현업", "적용", "실제", "실질", "바로", "즉시"],
    },
    # 시간/분량
    {
        "label": "교육 시간이나 분량 조정이 필요합니다.",
        "keywords": ["시간", "분량", "짧", "길", "부족", "많", "적었", "더 많", "늘려"],
    },
    # 반복 개최 요청
    {
        "label": "지속적인 교육 운영을 희망합니다.",
        "keywords": ["다음", "또", "재", "지속", "계속", "반복", "정기", "이어"],
    },
    # 환경/시설
    {
        "label": "교육 환경 및 시설이 쾌적했습니다.",
        "keywords": ["환경", "시설", "장소", "공간", "쾌적", "편리"],
    },
    # 개선 요청
    {
        "label": "교육 내용의 심화 또는 개선이 필요합니다.",
        "keywords": ["개선", "보완", "더", "추가", "심화", "발전", "향상", "업그레이드"],
    },
    # 동료/네트워크
    {
        "label": "동료와의 교류 및 네트워킹이 유익했습니다.",
        "keywords": ["동료", "네트워크", "교류", "소통", "함께", "팀", "협력"],
    },
]


def group_responses_rule_based(answers):
    """규칙 기반 공통응답 그룹핑"""
    groups = {}  # label → [answers]
    uncategorized = []
    
    for answer in answers:
        ans_lower = answer.lower()
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
            uncategorized.append(answer)
    
    result = []
    for lbl, ans_list in groups.items():
        result.append({
            "label": lbl,
            "answers": ans_list,
            "count": len(ans_list),
        })
    
    if uncategorized:
        result.append({
            "label": "기타 의견",
            "answers": uncategorized,
            "count": len(uncategorized),
        })
    
    # 가나다 오름차순 정렬
    result.sort(key=lambda x: x["label"])
    
    # 공통응답N 라벨 부여
    for i, item in enumerate(result):
        item["common_id"] = f"공통응답{i+1}"
    
    return result


def generate_grouping_prompt(open_ended_item):
    """Gemini AI에 보낼 공통응답 그룹핑 프롬프트"""
    answers = open_ended_item.get("answers", [])
    if not answers:
        return None
    
    answers_text = "\n".join([f"{i+1}. {a}" for i, a in enumerate(answers[:80])])
    
    prompt = f"""다음은 교육 만족도 설문의 주관식 응답입니다. 비슷한 의미의 응답들을 그룹으로 묶어주세요.

문항: {open_ended_item.get('label', '')}
응답 수: {len(answers)}건

응답 목록:
{answers_text}

규칙:
1. 비슷한 의미나 주제를 가진 응답들을 하나의 그룹으로 묶으세요
2. 각 그룹에 대표 문장(공통응답)을 만드세요 - 자연스러운 완전한 문장으로
3. 그룹은 가나다 오름차순으로 정렬하세요
4. 응답이 1건뿐인 경우도 별도 그룹으로 만들어도 됩니다
5. 응답 원문은 변경하지 마세요

출력 형식 (JSON):
{{
    "groups": [
        {{
            "label": "그룹 대표 공통응답 문장 (가나다 기준 정렬됨)",
            "count": 해당_응답_수,
            "answers": ["원문 응답 1", "원문 응답 2"]
        }}
    ]
}}

JSON만 출력하세요."""
    return prompt


async def group_responses_ai(open_ended_item, ai_engine):
    """AI 기반 공통응답 그룹핑"""
    if not ai_engine or not ai_engine.enabled:
        # 규칙 기반 폴백
        groups = group_responses_rule_based(open_ended_item.get("answers", []))
        return groups
    
    prompt = generate_grouping_prompt(open_ended_item)
    if not prompt:
        return []
    
    try:
        result = await ai_engine.generate_narrative(prompt)
        
        if isinstance(result, dict) and "groups" in result:
            groups = result["groups"]
            # 가나다 오름차순 보장
            groups.sort(key=lambda x: x.get("label", ""))
            # 공통응답N 라벨
            for i, g in enumerate(groups):
                g["common_id"] = f"공통응답{i+1}"
            return groups
    except Exception as e:
        print(f"AI 그룹핑 실패: {e}")
    
    # 폴백
    groups = group_responses_rule_based(open_ended_item.get("answers", []))
    return groups


async def process_all_open_ended(open_ended_list, ai_engine=None):
    """모든 주관식 문항에 공통응답 그룹핑 적용"""
    result = []
    for oe in open_ended_list:
        if not oe.get("answers"):
            result.append({**oe, "groups": []})
            continue
        
        groups = await group_responses_ai(oe, ai_engine)
        result.append({
            **oe,
            "groups": groups,
            "total_grouped": sum(g["count"] for g in groups),
        })
    
    return result
