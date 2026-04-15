"""
AI 추론 엔진 v2 — Gemini 2.0 Flash 기반
+ 전략적 내러티브 생성
+ 주관식 테마 요약
+ 슬라이드 구성안 추론
+ 자체 검증
"""
import json
import os

try:
    import google.generativeai as genai
    HAS_GENAI = True
except ImportError:
    HAS_GENAI = False


class AIEngine:
    def __init__(self, config):
        self.config = config
        api_key = config.get("gemini", {}).get("api_key", "") or os.environ.get("GEMINI_API_KEY", "")
        self.model_name = config.get("gemini", {}).get("model", "gemini-2.0-flash")
        self.enabled = bool(api_key) and HAS_GENAI
        
        if self.enabled:
            genai.configure(api_key=api_key)
            self.model = genai.GenerativeModel(self.model_name)
    
    # ═══════════════════════════════════════
    # 1. 전략적 내러티브 생성
    # ═══════════════════════════════════════
    
    async def generate_narrative(self, prompt):
        """Gemini로 전략적 비즈니스 내러티브 생성"""
        if not self.enabled:
            return self._fallback_narrative()
        
        try:
            response = self.model.generate_content(prompt)
            text = response.text.strip()
            if text.startswith("```"):
                text = text.split("\n", 1)[1].rsplit("```", 1)[0]
            return json.loads(text)
        except Exception as e:
            print(f"AI 내러티브 생성 실패: {e}")
            return self._fallback_narrative()
    
    async def summarize_qualitative(self, prompt):
        """Gemini로 주관식 테마 요약"""
        if not self.enabled:
            return self._fallback_qualitative()
        
        try:
            response = self.model.generate_content(prompt)
            text = response.text.strip()
            if text.startswith("```"):
                text = text.split("\n", 1)[1].rsplit("```", 1)[0]
            return json.loads(text)
        except Exception as e:
            print(f"AI 주관식 요약 실패: {e}")
            return self._fallback_qualitative()
    
    # ═══════════════════════════════════════
    # 2. 슬라이드 구성안 추론
    # ═══════════════════════════════════════
    
    async def infer_structure(self, summary, sample_patterns):
        """데이터 요약 + 샘플 패턴 → 최적 슬라이드 구성안 추론"""
        if not self.enabled:
            return self._fallback_inference(summary)
        
        prompt = f"""당신은 HRD 교육 만족도 결과보고서 전문가입니다.

아래 설문 데이터 요약을 보고, 가장 적합한 보고서 슬라이드 구성안을 JSON으로 출력하세요.

## 설문 데이터 요약
- 고객사: {summary['company']}
- 과정명: {summary['course_name']}
- 문항 수: {summary['total_questions']}개
- 카테고리: {summary['categories']}개 ({', '.join(summary['category_names'])})
- 모듈 유무: {'있음' if summary['has_modules'] else '없음'}
- 강사 수: {summary['num_instructors']}명
- 주관식 문항: {summary['open_ended_count']}개
- 응답인원: {summary['response_count']}명
- 전체 평균: {summary['overall_average']}점

## 출력 형식 (JSON)
{{
    "matched_pattern": "가장 유사한 샘플 패턴 이름",
    "similarity": 0.85,
    "reasoning": "이 데이터가 해당 패턴과 유사한 이유 설명",
    "recommended_slides": [
        {{"type": "exec_summary", "title": "Executive Summary"}},
        {{"type": "quant_general", "title": "정량 평가 결과"}},
        {{"type": "qual", "title": "정성 평가 결과"}}
    ],
    "additional_suggestions": []
}}

JSON만 출력하세요."""

        try:
            response = self.model.generate_content(prompt)
            text = response.text.strip()
            if text.startswith("```"):
                text = text.split("\n", 1)[1].rsplit("```", 1)[0]
            return json.loads(text)
        except Exception as e:
            print(f"AI 추론 실패: {e}")
            return self._fallback_inference(summary)
    
    # ═══════════════════════════════════════
    # 3. 자체 검증
    # ═══════════════════════════════════════
    
    async def review_output(self, data, preview_slides):
        """생성된 PPT 자체 검증"""
        checks = []
        
        if preview_slides:
            s1 = preview_slides[0]
            has_company = data["course_info"].get("company", "") in str(s1.get("texts", []))
            checks.append({
                "item": "표지 고객사명",
                "status": "pass" if has_company else "fail",
                "detail": "표지에 고객사명 포함" if has_company else "표지에 고객사명 누락"
            })
        
        total_qs = len(data["questions"])
        table_rows = sum(s.get("table_rows", 0) for s in preview_slides if "정량" in s.get("title", ""))
        checks.append({
            "item": "정량 문항 수",
            "status": "pass" if table_rows >= total_qs else "warn",
            "detail": f"데이터 {total_qs}문항, 표 {table_rows}행"
        })
        
        oe_count = len(data.get("open_ended", []))
        qual_groups = sum(s.get("group_count", 0) for s in preview_slides if "정성" in s.get("title", ""))
        checks.append({
            "item": "정성 평가 문항",
            "status": "pass" if qual_groups >= oe_count else "warn",
            "detail": f"주관식 {oe_count}개, 그룹 {qual_groups}개"
        })
        
        # AI 인사이트 존재 확인
        has_insight = bool(data.get("insights"))
        checks.append({
            "item": "데이터 인사이트",
            "status": "pass" if has_insight else "warn",
            "detail": "전략적 분석 포함" if has_insight else "원시 데이터만 포함"
        })
        
        # AI 내러티브 확인
        has_narrative = bool(data.get("narrative"))
        checks.append({
            "item": "AI 내러티브",
            "status": "pass" if has_narrative else "warn",
            "detail": "비즈니스 코멘트 포함" if has_narrative else "AI 코멘트 없음 (규칙 기반)"
        })
        
        passed = sum(1 for c in checks if c["status"] == "pass")
        return {"score": round(passed / len(checks) * 100), "checks": checks}
    
    # ═══════════════════════════════════════
    # 4. 수정 요청 해석
    # ═══════════════════════════════════════
    
    async def interpret_modification(self, user_request, summary, preview):
        if not self.enabled:
            return {"action": "manual", "detail": user_request}
        
        prompt = f"""사용자가 HRD 결과보고서 PPT에 대해 수정을 요청했습니다.

수정 요청: "{user_request}"
현재 보고서: {summary['company']} {summary['course_name']} ({len(preview)}슬라이드)

JSON으로 수정 액션을 출력하세요:
{{"action": "add_slide|modify_slide|delete_slide|modify_data", "target": "대상", "detail": "구체적 수정"}}"""
        try:
            response = self.model.generate_content(prompt)
            text = response.text.strip()
            if text.startswith("```"):
                text = text.split("\n", 1)[1].rsplit("```", 1)[0]
            return json.loads(text)
        except:
            return {"action": "manual", "detail": user_request}
    
    # ═══════════════════════════════════════
    # 폴백 (AI 없을 때)
    # ═══════════════════════════════════════
    
    def _fallback_narrative(self):
        return {
            "executive_summary": "",
            "strength_comment": "",
            "improvement_comment": "",
            "recommendation": "",
        }
    
    def _fallback_qualitative(self):
        return {
            "themes": [],
            "overall_sentiment": "",
        }
    
    def _fallback_inference(self, summary):
        slides = []
        cats = summary["category_names"]
        slides.append({"type": "exec_summary", "title": "Executive Summary", "categories": cats[:4]})
        if len(cats) > 4:
            slides.append({"type": "exec_summary_extra", "title": "Executive Summary", "categories": cats[4:]})
        if summary["has_modules"]:
            slides.append({"type": "exec_chart_module", "title": "모듈별 만족도"})
        if summary["num_instructors"] >= 2:
            slides.append({"type": "exec_chart_instructor", "title": "강사별 만족도"})
        slides.append({"type": "quant_general", "title": "정량 평가 결과"})
        if summary["has_modules"]:
            slides.append({"type": "quant_module", "title": "정량 평가 (모듈)"})
        if summary["num_instructors"] > 0:
            slides.append({"type": "quant_instructor", "title": "정량 평가 (강사)"})
        slides.append({"type": "qual", "title": "정성 평가", "group_count": summary["open_ended_count"]})
        return {
            "matched_pattern": "규칙 기반 (AI 미연결)",
            "similarity": 0,
            "reasoning": "규칙 기반으로 추론합니다.",
            "recommended_slides": slides,
            "additional_suggestions": [],
        }
