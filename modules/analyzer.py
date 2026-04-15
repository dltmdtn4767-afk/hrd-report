"""
analyzer.py — 데이터 해석 모듈
Raw 평균 → Strategic Insight 변환

- score_gap: 기준(4.0)과의 차이
- performance_tier: 상/중/하
- top/bottom 카테고리 식별
- 트렌드 방향 (이전 데이터 있으면)
"""

BENCHMARK = 4.0  # 업계 기준점


def analyze_data(data, prev_data=None):
    """데이터에 분석 컬럼 추가 + 전략적 요약 생성"""
    
    questions = data["questions"]
    categories = data["categories"]
    overall = data.get("overall_average", 0)
    
    # ═══ 문항별 분석 ═══
    all_avgs = [q["avg"] for q in questions if q["avg"] > 0]
    q_mean = sum(all_avgs) / len(all_avgs) if all_avgs else BENCHMARK
    
    for q in questions:
        q["score_gap"] = round(q["avg"] - BENCHMARK, 2)
        q["internal_gap"] = round(q["avg"] - q_mean, 2)
        
        if q["avg"] >= 4.5:
            q["tier"] = "excellent"
        elif q["avg"] >= 4.0:
            q["tier"] = "good"
        elif q["avg"] >= 3.5:
            q["tier"] = "moderate"
        else:
            q["tier"] = "needs_improvement"
        
        # 트렌드 (이전 데이터 있으면)
        if prev_data:
            prev_q = next((pq for pq in prev_data.get("questions", []) if pq["id"] == q["id"]), None)
            if prev_q:
                diff = q["avg"] - prev_q["avg"]
                q["trend"] = "up" if diff > 0.1 else "down" if diff < -0.1 else "stable"
                q["trend_delta"] = round(diff, 2)
            else:
                q["trend"] = "new"
                q["trend_delta"] = 0
        else:
            q["trend"] = "N/A"
            q["trend_delta"] = 0
    
    # ═══ 카테고리별 분석 ═══
    for cat in categories:
        cat["score_gap"] = round(cat["avg"] - BENCHMARK, 2)
        cat["internal_gap"] = round(cat["avg"] - overall, 2)
        cat["tier"] = _tier(cat["avg"])
        
        # 카테고리 내 최고/최저 문항
        cat_qs = sorted(cat["questions"], key=lambda q: q["avg"], reverse=True)
        cat["best_question"] = cat_qs[0] if cat_qs else None
        cat["worst_question"] = cat_qs[-1] if len(cat_qs) > 1 else None
    
    # ═══ 전략적 식별 ═══
    sorted_cats = sorted(categories, key=lambda c: c["avg"], reverse=True)
    
    insights = {
        "overall_average": overall,
        "overall_gap": round(overall - BENCHMARK, 2),
        "overall_tier": _tier(overall),
        "top_categories": sorted_cats[:3],
        "bottom_categories": sorted_cats[-1:] if len(sorted_cats) > 1 else [],
        "above_benchmark_count": sum(1 for c in categories if c["avg"] >= BENCHMARK),
        "total_categories": len(categories),
        "response_count": data.get("response_count", 0),
        "std_dev": _std_dev([c["avg"] for c in categories]),
        "consistency": "높음" if _std_dev([c["avg"] for c in categories]) < 0.3 else "보통" if _std_dev([c["avg"] for c in categories]) < 0.5 else "낮음",
    }
    
    # ═══ 주관식 키워드 분석 ═══
    if data.get("open_ended"):
        pos_keywords = ["좋", "만족", "도움", "유익", "추천", "감사", "최고", "훌륭", "적극"]
        neg_keywords = ["아쉬", "부족", "개선", "짧", "길", "어려", "불만", "힘들"]
        
        all_answers = []
        for oe in data["open_ended"]:
            all_answers.extend(oe.get("answers", []))
        
        pos_count = sum(1 for a in all_answers if any(kw in a for kw in pos_keywords))
        neg_count = sum(1 for a in all_answers if any(kw in a for kw in neg_keywords))
        
        insights["qualitative"] = {
            "total_responses": len(all_answers),
            "positive_ratio": round(pos_count / max(len(all_answers), 1) * 100),
            "negative_ratio": round(neg_count / max(len(all_answers), 1) * 100),
        }
    
    data["insights"] = insights
    return data


def generate_narrative_prompt(data):
    """Gemini AI에 보낼 전략적 내러티브 프롬프트 생성"""
    insights = data.get("insights", {})
    course = data.get("course_info", {})
    
    top = insights.get("top_categories", [])
    bottom = insights.get("bottom_categories", [])
    
    top_text = ", ".join([f"{c['name']}({c['avg']:.2f}점)" for c in top])
    bottom_text = ", ".join([f"{c['name']}({c['avg']:.2f}점)" for c in bottom])
    
    prompt = f"""당신은 HRD 교육 컨설팅 전문가입니다.
아래 교육 만족도 데이터를 바탕으로 전문적인 비즈니스 코멘트를 작성하세요.

## 교육 정보
- 고객사: {course.get('company', '')}
- 과정명: {course.get('course_name', '')}
- 응답인원: {insights.get('response_count', 0)}명
- 전체 평균: {insights.get('overall_average', 0):.2f}점 / 5점
- 기준점(4.0) 대비: {'+' if insights.get('overall_gap', 0) >= 0 else ''}{insights.get('overall_gap', 0):.2f}
- 카테고리 균일도: {insights.get('consistency', '')}

## 주요 성과 (상위 카테고리)
{top_text}

## 개선 영역 (하위 카테고리)
{bottom_text}

## 출력 형식 (JSON)
{{
    "executive_summary": "3문장 분량의 전문 비즈니스 요약. 강점은 '성과'로, 약점은 '향후 개선 기회'로 표현. 높임말 사용. 구체적 점수 포함.",
    "strength_comment": "상위 카테고리에 대한 1-2문장 코멘트",
    "improvement_comment": "하위 카테고리에 대한 건설적 1-2문장 코멘트. '미흡'이 아닌 '성장 잠재력' 프레이밍.",
    "recommendation": "향후 교육 운영에 대한 1문장 제안"
}}

JSON만 출력. 설명 없이."""
    
    return prompt


def generate_qualitative_prompt(open_ended):
    """주관식 응답 → AI 테마 요약 프롬프트"""
    all_answers = []
    for oe in open_ended:
        for a in oe.get("answers", []):
            if len(a) > 3:
                all_answers.append(a)
    
    answers_text = "\n".join([f"- {a}" for a in all_answers[:50]])  # 최대 50개
    
    prompt = f"""아래 교육 만족도 주관식 응답들을 분석하여 핵심 테마 3개로 요약하세요.

## 주관식 응답 ({len(all_answers)}건)
{answers_text}

## 출력 형식 (JSON)
{{
    "themes": [
        {{"title": "테마 제목", "summary": "2문장 요약", "sentiment": "positive/negative/neutral", "count": 대략적 관련 응답 수}},
        {{"title": "테마 제목", "summary": "2문장 요약", "sentiment": "positive/negative/neutral", "count": 대략적 관련 응답 수}},
        {{"title": "테마 제목", "summary": "2문장 요약", "sentiment": "positive/negative/neutral", "count": 대략적 관련 응답 수}}
    ],
    "overall_sentiment": "전반적 감성 한 줄 요약"
}}

JSON만 출력."""
    
    return prompt


def _tier(score):
    if score >= 4.5: return "excellent"
    elif score >= 4.0: return "good"
    elif score >= 3.5: return "moderate"
    else: return "needs_improvement"


def _std_dev(values):
    if len(values) < 2:
        return 0
    mean = sum(values) / len(values)
    variance = sum((x - mean) ** 2 for x in values) / len(values)
    return round(variance ** 0.5, 3)
