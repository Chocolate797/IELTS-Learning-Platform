from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import Optional, List, Dict, Any
from datetime import date, datetime
import openpyxl
import csv
import io
import json
import os
from openai import OpenAI

app = FastAPI(title="FinanceAI Boss API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY", ""))

class FinancialRecord(BaseModel):
    id: str
    date: str
    type: str
    category: str
    amount: float
    description: str

class AnalysisRequest(BaseModel):
    records: List[FinancialRecord]
    date_range: Optional[Dict[str, str]] = None
    granularity: str = "monthly"

class ChatRequest(BaseModel):
    message: str
    context: Optional[Dict[str, Any]] = None

class KPIData(BaseModel):
    revenue: float
    cost: float
    profit: float
    profit_margin: float
    revenue_trend: float
    cost_trend: float
    profit_trend: float

@app.get("/")
async def root():
    return {"message": "FinanceAI Boss API is running", "version": "1.0.0"}

@app.post("/api/v1/analyze")
async def analyze_financial_data(request: AnalysisRequest):
    records = request.records
    if not records:
        raise HTTPException(status_code=400, detail="No records provided")

    income_records = [r for r in records if r.type == "income"]
    expense_records = [r for r in records if r.type == "expense"]

    total_income = sum(r.amount for r in income_records)
    total_expense = sum(r.amount for r in expense_records)
    profit = total_income - total_expense
    profit_margin = (profit / total_income * 100) if total_income > 0 else 0

    monthly_data = {}
    for r in records:
        month = r.date[:7]
        if month not in monthly_data:
            monthly_data[month] = {"income": 0, "expense": 0}
        if r.type == "income":
            monthly_data[month]["income"] += r.amount
        else:
            monthly_data[month]["expense"] += r.amount

    sorted_months = sorted(monthly_data.keys())
    trends = []
    for month in sorted_months:
        income = monthly_data[month]["income"]
        expense = monthly_data[month]["expense"]
        trends.append({
            "month": month,
            "income": income,
            "expense": expense,
            "profit": income - expense
        })

    category_breakdown = {}
    for r in expense_records:
        if r.category not in category_breakdown:
            category_breakdown[r.category] = 0
        category_breakdown[r.category] += r.amount

    alerts = []
    if len(sorted_months) >= 2:
        current_month = sorted_months[-1]
        prev_month = sorted_months[-2]
        current_expense = monthly_data[current_month]["expense"]
        prev_expense = monthly_data[prev_month]["expense"]
        if prev_expense > 0:
            expense_change = (current_expense - prev_expense) / prev_expense * 100
            if expense_change > 20:
                alerts.append({
                    "type": "cost_spike",
                    "severity": "high",
                    "message": f"支出环比增长 {expense_change:.1f}%，需关注",
                    "details": {
                        "current": current_expense,
                        "previous": prev_expense,
                        "change_percent": expense_change
                    }
                })

    if profit_margin < 10:
        alerts.append({
            "type": "low_margin",
            "severity": "medium",
            "message": f"利润率仅 {profit_margin:.1f}%，低于健康水平",
            "details": {"margin": profit_margin}
        })

    return {
        "kpi": {
            "revenue": total_income,
            "cost": total_expense,
            "profit": profit,
            "profit_margin": profit_margin,
            "revenue_trend": 0,
            "cost_trend": 0,
            "profit_trend": 0
        },
        "trends": trends,
        "category_breakdown": category_breakdown,
        "alerts": alerts
    }

@app.post("/api/v1/chat")
async def chat_with_ai(request: ChatRequest):
    if not client.api_key:
        raise HTTPException(status_code=500, detail="OpenAI API key not configured")

    system_prompt = """你是一位专业、友善的中小企业财务顾问。用户是企业老板，不具备专业财务知识。

请遵循以下原则：
1. 用通俗易懂的语言解释财务问题，避免专业术语
2. 结合具体数字给出分析
3. 给出明确、可执行的建议
4. 如有风险，明确指出
5. 回答要简洁有力，不要冗长

财务数据上下文（如果有）：
{context}

请根据上下文回答用户的问题。"""

    context_str = ""
    if request.context:
        kpi = request.context.get("kpi", {})
        if kpi:
            context_str += f"""
当前财务概况：
- 总收入：¥{kpi.get('revenue', 0):,.2f}
- 总支出：¥{kpi.get('cost', 0):,.2f}
- 净利润：¥{kpi.get('profit', 0):,.2f}
- 利润率：{kpi.get('profit_margin', 0):.1f}%
"""
        alerts = request.context.get("alerts", [])
        if alerts:
            context_str += "\n预警信息：\n"
            for alert in alerts:
                context_str += f"- {alert.get('message', '')}\n"

    full_prompt = system_prompt.format(context=context_str) + f"\n\n用户问题：{request.message}"

    try:
        response = client.chat.completions.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": system_prompt.format(context=context_str)},
                {"role": "user", "content": request.message}
            ],
            temperature=0.7,
            max_tokens=1000
        )

        reply = response.choices[0].message.content

        return {
            "reply": reply,
            "suggestions": [
                "分析成本变化原因",
                "查看现金流状况",
                "获取决策建议"
            ]
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/v1/health")
async def health_check():
    return {
        "status": "healthy",
        "timestamp": datetime.now().isoformat(),
        "openai_configured": bool(client.api_key)
    }

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
