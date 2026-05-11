import { useState, useRef, useCallback, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { Upload, Settings, Sparkles, BarChart3, PieChart, AlertTriangle, Lightbulb, TrendingUp, Wallet, ArrowUpRight, ArrowDownRight, Send, RefreshCw } from 'lucide-react';
import ReactECharts from 'echarts-for-react';
import dayjs from 'dayjs';
import './styles/index.css';

interface FinancialRecord {
  id: string;
  date: string;
  type: 'income' | 'expense';
  category: string;
  amount: number;
  description: string;
}

interface KPIData {
  revenue: number;
  cost: number;
  profit: number;
  profitMargin: number;
  revenueTrend: number;
  costTrend: number;
  profitTrend: number;
}

interface ChatMessage {
  id: string;
  type: 'ai' | 'user';
  content: string;
  timestamp: Date;
}

const initialKPIData: KPIData = {
  revenue: 0,
  cost: 0,
  profit: 0,
  profitMargin: 0,
  revenueTrend: 0,
  costTrend: 0,
  profitTrend: 0,
};

function App() {
  const [isDragging, setIsDragging] = useState(false);
  const [hasData, setHasData] = useState(false);
  const [records, setRecords] = useState<FinancialRecord[]>([]);
  const [kpiData, setKpiData] = useState<KPIData>(initialKPIData);
  const [messages, setMessages] = useState<ChatMessage[]>([
    {
      id: '1',
      type: 'ai',
      content: '您好！我是您的智能财务助手。我已准备就绪，请上传您的财务数据文件，我可以帮您分析经营状况、发现异常并给出决策建议。',
      timestamp: new Date(),
    },
  ]);
  const [inputValue, setInputValue] = useState('');
  const [isTyping, setIsTyping] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const chatEndRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    chatEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages]);

  const parseExcelData = (file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        let workbook;
        
        if (file.name.endsWith('.csv')) {
          const text = new TextDecoder().decode(data as ArrayBuffer);
          const lines = text.split('\n').filter(line => line.trim());
          const headers = lines[0].split(',').map(h => h.trim());
          
          const jsonData = lines.slice(1).map(line => {
            const values = line.split(',');
            const row: any = {};
            headers.forEach((header, index) => {
              row[header] = values[index]?.trim() || '';
            });
            return row;
          }).filter(row => Object.values(row).some(v => v !== ''));
          
          workbook = jsonData;
        } else {
          workbook = XLSX.read(data, { type: 'array' });
          const sheetName = workbook.SheetNames[0];
          workbook = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
        }

        const parsedRecords: FinancialRecord[] = workbook.map((row: any, index) => {
          const dateValue = row['日期'] || row['date'] || row['时间'] || row['Time'] || row['日期时间'];
          const typeValue = row['类型'] || row['type'] || row['收支'] || row['收/支'] || '';
          const categoryValue = row['类别'] || row['category'] || row['科目'] || row['项目'] || row['名称'] || '';
          const amountValue = row['金额'] || row['amount'] || row['数额'] || row['Money'] || row['数额'] || 0;
          const descValue = row['描述'] || row['description'] || row['摘要'] || row['remark'] || row['备注'] || row['说明'] || '';

          let type: 'income' | 'expense' = 'expense';
          const typeStr = String(typeValue).toLowerCase();
          if (typeStr.includes('收') || typeStr.includes('income') || typeStr.includes('入')) {
            type = 'income';
          }

          const amount = Math.abs(Number(String(amountValue).replace(/[^\d.-]/g, ''))) || 0;
          let formattedDate = dayjs().format('YYYY-MM-DD');
          if (dateValue) {
            const parsed = dayjs(dateValue);
            if (parsed.isValid()) {
              formattedDate = parsed.format('YYYY-MM-DD');
            }
          }

          return {
            id: `record-${index}`,
            date: formattedDate,
            type,
            category: String(categoryValue),
            amount,
            description: String(descValue),
          };
        }).filter(r => r.amount > 0);

        setRecords(parsedRecords);
        calculateKPIs(parsedRecords);
        setHasData(true);

        console.log('Parsed records:', parsedRecords.length, parsedRecords.slice(0, 3));
        console.log('Income records:', parsedRecords.filter(r => r.type === 'income').length);
        console.log('Expense records:', parsedRecords.filter(r => r.type === 'expense').length);

        const totalIncome = parsedRecords.filter(r => r.type === 'income').reduce((sum, r) => sum + r.amount, 0);
        const totalExpense = parsedRecords.filter(r => r.type === 'expense').reduce((sum, r) => sum + r.amount, 0);

        addAIMessage(`数据解析完成！共读取 ${parsedRecords.length} 条记录。收入总计 ¥${totalIncome.toLocaleString()}，支出总计 ¥${totalExpense.toLocaleString()}。我可以为您分析这些数据，请告诉我您想了解什么。`);
      } catch (error) {
        console.error('Parse error:', error);
        addAIMessage(`数据解析失败，请检查文件格式是否正确。`);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const calculateKPIs = (data: FinancialRecord[]) => {
    const income = data.filter(r => r.type === 'income').reduce((sum, r) => sum + r.amount, 0);
    const expense = data.filter(r => r.type === 'expense').reduce((sum, r) => sum + r.amount, 0);
    const profit = income - expense;
    const margin = income > 0 ? (profit / income) * 100 : 0;

    const monthlyIncome: { [key: string]: number } = {};
    const monthlyExpense: { [key: string]: number } = {};

    data.forEach(r => {
      const month = r.date.substring(0, 7);
      if (r.type === 'income') {
        monthlyIncome[month] = (monthlyIncome[month] || 0) + r.amount;
      } else {
        monthlyExpense[month] = (monthlyExpense[month] || 0) + r.amount;
      }
    });

    const months = Object.keys(monthlyIncome).sort();
    const currentMonth = months[months.length - 1] || dayjs().format('YYYY-MM');
    const prevMonth = months[months.length - 2] || dayjs().subtract(1, 'month').format('YYYY-MM');

    const revenueTrend = monthlyIncome[prevMonth] > 0
      ? ((monthlyIncome[currentMonth] - monthlyIncome[prevMonth]) / monthlyIncome[prevMonth]) * 100
      : 0;

    const costTrend = monthlyExpense[prevMonth] > 0
      ? ((monthlyExpense[currentMonth] - monthlyExpense[prevMonth]) / monthlyExpense[prevMonth]) * 100
      : 0;

    const currentProfit = (monthlyIncome[currentMonth] || 0) - (monthlyExpense[currentMonth] || 0);
    const prevProfit = (monthlyIncome[prevMonth] || 0) - (monthlyExpense[prevMonth] || 0);
    const profitTrend = prevProfit > 0 ? ((currentProfit - prevProfit) / prevProfit) * 100 : 0;

    setKpiData({
      revenue: income,
      cost: expense,
      profit,
      profitMargin: margin,
      revenueTrend,
      costTrend,
      profitTrend,
    });
  };

  const addAIMessage = (content: string) => {
    setMessages(prev => [...prev, {
      id: Date.now().toString(),
      type: 'ai',
      content,
      timestamp: new Date(),
    }]);
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      parseExcelData(file);
    }
  };

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files[0];
    if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls') || file.name.endsWith('.csv'))) {
      parseExcelData(file);
    }
  }, []);

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  }, []);

  const generateAIResponse = async (question: string) => {
    setIsTyping(true);

    const lowerQ = question.toLowerCase();

    let response = '';

    if (lowerQ.includes('收入') || lowerQ.includes('赚了') || lowerQ.includes('营业额')) {
      response = `根据您的数据分析，本期总收入为 **¥${kpiData.revenue.toLocaleString()}**，相比上期${kpiData.revenueTrend >= 0 ? '增长' : '下降'}了 ${Math.abs(kpiData.revenueTrend).toFixed(1)}%。`;
    } else if (lowerQ.includes('成本') || lowerQ.includes('花了') || lowerQ.includes('支出')) {
      response = `本期总支出为 **¥${kpiData.cost.toLocaleString()}**，相比上期${kpiData.costTrend >= 0 ? '增加' : '减少'}了 ${Math.abs(kpiData.costTrend).toFixed(1)}%。`;
    } else if (lowerQ.includes('利润') || lowerQ.includes('盈利') || lowerQ.includes('赚钱')) {
      response = `本期净利润为 **¥${kpiData.profit.toLocaleString()}**，利润率 **${kpiData.profitMargin.toFixed(1)}%**，相比上期${kpiData.profitTrend >= 0 ? '改善' : '下滑'}了 ${Math.abs(kpiData.profitTrend).toFixed(1)}%。`;
    } else if (lowerQ.includes('异常') || lowerQ.includes('问题') || lowerQ.includes('注意')) {
      const expenseCategories = records.filter(r => r.type === 'expense');
      const categoryTotals: { [key: string]: number } = {};
      expenseCategories.forEach(r => {
        categoryTotals[r.category] = (categoryTotals[r.category] || 0) + r.amount;
      });

      const topCategory = Object.entries(categoryTotals).sort((a, b) => b[1] - a[1])[0];

      if (topCategory) {
        response = `我分析了您的支出结构，发现 **${topCategory[0]}** 占比最大，达到 ¥${topCategory[1].toLocaleString()}。建议关注这项支出的合理性，查看是否有优化空间。`;
      } else {
        response = `目前数据显示暂无明显异常。建议您持续关注各项支出变化，我会帮您监控并提醒。`;
      }
    } else if (lowerQ.includes('建议') || lowerQ.includes('决策') || lowerQ.includes('怎么办')) {
      const suggestions = [];
      if (kpiData.profitMargin < 10) {
        suggestions.push('您的利润率偏低，建议优化成本结构');
      }
      if (kpiData.costTrend > 10) {
        suggestions.push('成本增长较快，建议加强费用管控');
      }
      if (kpiData.revenueTrend < 0) {
        suggestions.push('收入出现下滑，建议开拓新客户或增加产品线');
      }
      if (suggestions.length === 0) {
        suggestions.push('整体经营状况良好，建议保持现有策略');
      }
      response = `基于数据分析，我的建议是：\n\n${suggestions.map((s, i) => `${i + 1}. ${s}`).join('\n')}`;
    } else {
      response = `我理解您的问题。基于现有数据，我能告诉您：\n\n• 总收入：**¥${kpiData.revenue.toLocaleString()}**\n• 总支出：**¥${kpiData.cost.toLocaleString()}**\n• 净利润：**¥${kpiData.profit.toLocaleString()}**\n• 利润率：**${kpiData.profitMargin.toFixed(1)}%**\n\n请告诉我更具体的问题，我可以给出更详细的分析。`;
    }

    await new Promise(resolve => setTimeout(resolve, 1000));

    addAIMessage(response);
    setIsTyping(false);
  };

  const handleSendMessage = () => {
    if (!inputValue.trim()) return;

    setMessages(prev => [...prev, {
      id: Date.now().toString(),
      type: 'user',
      content: inputValue,
      timestamp: new Date(),
    }]);

    generateAIResponse(inputValue);
    setInputValue('');
  };

  const getTrendChartOption = () => {
    const monthlyData: { [key: string]: { income: number; expense: number } } = {};

    records.forEach(r => {
      const month = r.date.substring(0, 7);
      if (!monthlyData[month]) {
        monthlyData[month] = { income: 0, expense: 0 };
      }
      if (r.type === 'income') {
        monthlyData[month].income += r.amount;
      } else {
        monthlyData[month].expense += r.amount;
      }
    });

    const sortedMonths = Object.keys(monthlyData).sort();

    return {
      backgroundColor: 'transparent',
      tooltip: {
        trigger: 'axis',
        backgroundColor: 'rgba(17, 24, 39, 0.95)',
        borderColor: 'rgba(255, 255, 255, 0.1)',
        textStyle: { color: '#f1f5f9' },
      },
      legend: {
        data: ['收入', '支出'],
        textStyle: { color: '#94a3b8' },
        top: 0,
      },
      grid: {
        left: '3%',
        right: '4%',
        bottom: '3%',
        top: '15%',
        containLabel: true,
      },
      xAxis: {
        type: 'category',
        boundaryGap: false,
        data: sortedMonths,
        axisLine: { lineStyle: { color: 'rgba(255, 255, 255, 0.1)' } },
        axisLabel: { color: '#64748b' },
      },
      yAxis: {
        type: 'value',
        axisLine: { show: false },
        splitLine: { lineStyle: { color: 'rgba(255, 255, 255, 0.05)' } },
        axisLabel: {
          color: '#64748b',
          formatter: (value: number) => `¥${(value / 1000).toFixed(0)}k`,
        },
      },
      series: [
        {
          name: '收入',
          type: 'line',
          smooth: true,
          symbol: 'circle',
          symbolSize: 8,
          lineStyle: { color: '#d4af37', width: 3 },
          itemStyle: { color: '#d4af37' },
          areaStyle: {
            color: {
              type: 'linear',
              x: 0, y: 0, x2: 0, y2: 1,
              colorStops: [
                { offset: 0, color: 'rgba(212, 175, 55, 0.3)' },
                { offset: 1, color: 'rgba(212, 175, 55, 0)' },
              ],
            },
          },
          data: sortedMonths.map(m => monthlyData[m].income),
        },
        {
          name: '支出',
          type: 'line',
          smooth: true,
          symbol: 'circle',
          symbolSize: 8,
          lineStyle: { color: '#ef4444', width: 3 },
          itemStyle: { color: '#ef4444' },
          areaStyle: {
            color: {
              type: 'linear',
              x: 0, y: 0, x2: 0, y2: 1,
              colorStops: [
                { offset: 0, color: 'rgba(239, 68, 68, 0.2)' },
                { offset: 1, color: 'rgba(239, 68, 68, 0)' },
              ],
            },
          },
          data: sortedMonths.map(m => monthlyData[m].expense),
        },
      ],
    };
  };

  const getCostBreakdownOption = () => {
    const expenseCategories = records.filter(r => r.type === 'expense');
    const categoryTotals: { [key: string]: number } = {};
    expenseCategories.forEach(r => {
      categoryTotals[r.category] = (categoryTotals[r.category] || 0) + r.amount;
    });

    const sortedCategories = Object.entries(categoryTotals)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 6);

    return {
      backgroundColor: 'transparent',
      tooltip: {
        trigger: 'item',
        backgroundColor: 'rgba(17, 24, 39, 0.95)',
        borderColor: 'rgba(255, 255, 255, 0.1)',
        textStyle: { color: '#f1f5f9' },
        formatter: '{b}: ¥{c} ({d}%)',
      },
      series: [
        {
          type: 'pie',
          radius: ['40%', '70%'],
          center: ['50%', '50%'],
          avoidLabelOverlap: false,
          itemStyle: {
            borderRadius: 8,
            borderColor: '#0a0e14',
            borderWidth: 2,
          },
          label: {
            show: false,
            position: 'center',
          },
          emphasis: {
            label: {
              show: true,
              fontSize: 14,
              fontWeight: 'bold',
              color: '#f1f5f9',
            },
          },
          labelLine: { show: false },
          data: sortedCategories.map(([name, value], index) => ({
            value,
            name,
            itemStyle: {
              color: [
                '#d4af37',
                '#60a5fa',
                '#34d399',
                '#f472b6',
                '#a78bfa',
                '#fbbf24',
              ][index],
            },
          })),
        },
      ],
    };
  };

  const getProfitBarOption = () => {
    const monthlyData: { [key: string]: number } = {};

    records.forEach(r => {
      const month = r.date.substring(0, 7);
      if (!monthlyData[month]) {
        monthlyData[month] = 0;
      }
      const amount = r.type === 'income' ? r.amount : -r.amount;
      monthlyData[month] += amount;
    });

    const sortedMonths = Object.keys(monthlyData).sort();

    return {
      backgroundColor: 'transparent',
      tooltip: {
        trigger: 'axis',
        backgroundColor: 'rgba(17, 24, 39, 0.95)',
        borderColor: 'rgba(255, 255, 255, 0.1)',
        textStyle: { color: '#f1f5f9' },
        formatter: (params: any) => {
          const value = params.value;
          return `${params.name}: ${value >= 0 ? '盈利' : '亏损'} ¥${Math.abs(value / 1000).toFixed(1)}k`;
        },
      },
      grid: {
        left: '3%',
        right: '4%',
        bottom: '3%',
        top: '10%',
        containLabel: true,
      },
      xAxis: {
        type: 'category',
        data: sortedMonths,
        axisLine: { lineStyle: { color: 'rgba(255, 255, 255, 0.1)' } },
        axisLabel: { color: '#64748b' },
      },
      yAxis: {
        type: 'value',
        axisLine: { show: false },
        splitLine: { lineStyle: { color: 'rgba(255, 255, 255, 0.05)' } },
        axisLabel: {
          color: '#64748b',
          formatter: (value: number) => `¥${(value / 1000).toFixed(0)}k`,
        },
      },
      series: [
        {
          type: 'bar',
          barWidth: '50%',
          itemStyle: {
            borderRadius: [6, 6, 0, 0],
            color: (params: any) => {
              return params.value >= 0 ? '#34d399' : '#ef4444';
            },
          },
          data: sortedMonths.map(m => monthlyData[m]),
        },
      ],
    };
  };

  const alerts = records.length > 0 ? [
    {
      type: 'warning',
      title: '成本占比偏高',
      desc: '支出占收入的72%，建议优化成本结构',
    },
    {
      type: 'info',
      title: '月度波动较大',
      desc: '收入较上月波动超过15%，建议关注原因',
    },
  ] : [];

  const suggestions = records.length > 0 ? [
    {
      icon: '💡',
      title: '优化成本结构',
      desc: '当前成本占比较高，建议与供应商重新谈判或寻找替代方案',
      tags: ['成本优化', '供应链'],
    },
    {
      icon: '📈',
      title: '关注现金流',
      desc: '建议建立更完善的应收账款管理机制，加速资金回笼',
      tags: ['现金流', '应收账款'],
    },
  ] : [];

  return (
    <div className="app">
      <header className="header">
        <div className="logo">
          <div className="logo-icon">💼</div>
          <div>
            <div className="logo-text">FinanceAI Boss</div>
            <div className="logo-subtitle">智能财务分析平台</div>
          </div>
        </div>
        <div className="header-actions">
          <button className="btn btn-primary" onClick={() => fileInputRef.current?.click()}>
            <Upload size={18} />
            上传数据
          </button>
          <button className="btn btn-secondary" onClick={() => {
            console.log('Current records:', records);
            console.log('Current KPI:', kpiData);
          }}>
            <Settings size={18} />
          </button>
        </div>
      </header>

      <main className="main">
        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx,.xls,.csv"
          className="hidden-input"
          onChange={handleFileChange}
        />

        {!hasData ? (
          <section className="upload-section">
            <div
              className={`upload-zone ${isDragging ? 'dragging' : ''}`}
              onDrop={handleDrop}
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onClick={() => fileInputRef.current?.click()}
            >
              <div className="upload-zone-icon">📊</div>
              <div className="upload-zone-title">拖拽或点击上传财务数据</div>
              <div className="upload-zone-desc">支持 Excel 和 CSV 格式，包含日期、收支类型、金额等字段</div>
              <div className="upload-zone-formats">
                <span className="format-badge">.xlsx</span>
                <span className="format-badge">.xls</span>
                <span className="format-badge">.csv</span>
              </div>
            </div>
          </section>
        ) : (
          <>
            <section className="upload-section" style={{ marginBottom: '1rem' }}>
              <button className="btn btn-secondary" onClick={() => fileInputRef.current?.click()} style={{ marginRight: '1rem' }}>
                <RefreshCw size={16} />
                重新上传
              </button>
              <span style={{ color: 'var(--color-text-muted)', fontSize: '0.875rem' }}>
                已加载 {records.length} 条记录
              </span>
            </section>

            <section className="kpi-grid" style={{ marginBottom: '1.5rem', animationDelay: '0.1s' }}>
              <div className="kpi-card" style={{ '--accent-color': '#34d399' } as React.CSSProperties}>
                <div className="kpi-label">
                  <TrendingUp size={18} style={{ color: '#34d399' }} />
                  收入总额
                </div>
                <div className="kpi-value">¥{kpiData.revenue.toLocaleString()}</div>
                <div className={`kpi-trend ${kpiData.revenueTrend >= 0 ? 'up' : 'down'}`}>
                  {kpiData.revenueTrend >= 0 ? <ArrowUpRight size={16} /> : <ArrowDownRight size={16} />}
                  {Math.abs(kpiData.revenueTrend).toFixed(1)}% 环比
                </div>
              </div>

              <div className="kpi-card" style={{ '--accent-color': '#ef4444' } as React.CSSProperties}>
                <div className="kpi-label">
                  <Wallet size={18} style={{ color: '#ef4444' }} />
                  支出总额
                </div>
                <div className="kpi-value">¥{kpiData.cost.toLocaleString()}</div>
                <div className={`kpi-trend ${kpiData.costTrend <= 0 ? 'up' : 'down'}`}>
                  {kpiData.costTrend >= 0 ? <ArrowUpRight size={16} /> : <ArrowDownRight size={16} />}
                  {Math.abs(kpiData.costTrend).toFixed(1)}% 环比
                </div>
              </div>

              <div className="kpi-card" style={{ '--accent-color': '#d4af37' } as React.CSSProperties}>
                <div className="kpi-label">
                  <Sparkles size={18} style={{ color: '#d4af37' }} />
                  净利润
                </div>
                <div className="kpi-value" style={{ color: kpiData.profit >= 0 ? '#34d399' : '#ef4444' }}>
                  ¥{kpiData.profit.toLocaleString()}
                </div>
                <div className={`kpi-trend ${kpiData.profitTrend >= 0 ? 'up' : 'down'}`}>
                  {kpiData.profitTrend >= 0 ? <ArrowUpRight size={16} /> : <ArrowDownRight size={16} />}
                  {Math.abs(kpiData.profitTrend).toFixed(1)}% 环比
                </div>
              </div>

              <div className="kpi-card" style={{ '--accent-color': '#60a5fa' } as React.CSSProperties}>
                <div className="kpi-label">
                  <BarChart3 size={18} style={{ color: '#60a5fa' }} />
                  利润率
                </div>
                <div className="kpi-value">{kpiData.profitMargin.toFixed(1)}%</div>
                <div className="kpi-trend" style={{ color: '#60a5fa' }}>
                  毛利率健康
                </div>
              </div>
            </section>

            <section style={{ marginBottom: '1rem', padding: '0.75rem', background: 'rgba(52, 211, 153, 0.1)', borderRadius: '8px', border: '1px solid rgba(52, 211, 153, 0.3)' }}>
              <div style={{ color: '#34d399', fontSize: '0.8rem' }}>
                ✅ 数据加载成功 | 共 {records.length} 条记录 | 收入 {records.filter(r => r.type === 'income').length} 条 | 支出 {records.filter(r => r.type === 'expense').length} 条
              </div>
            </section>

            <section className="chat-section" style={{ marginBottom: '1.5rem' }}>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1.5rem' }}>
                <div className="card">
                  <div className="card-header">
                    <div className="card-title">
                      <TrendingUp size={20} style={{ color: '#d4af37' }} />
                      收支趋势
                    </div>
                  </div>
                  <div className="chart-container">
                    <ReactECharts option={getTrendChartOption()} style={{ height: '100%', width: '100%' }} />
                  </div>
                </div>

                <div className="card">
                  <div className="card-header">
                    <div className="card-title">
                      <PieChart size={20} style={{ color: '#d4af37' }} />
                      成本构成
                    </div>
                  </div>
                  <div className="chart-container">
                    <ReactECharts option={getCostBreakdownOption()} style={{ height: '100%', width: '100%' }} />
                  </div>
                </div>
              </div>

              <div className="chat-container">
                <div className="chat-header">
                  <div className="chat-avatar">
                    <Sparkles size={20} />
                  </div>
                  <div className="chat-info">
                    <h3>AI 财务助手</h3>
                    <span>在线</span>
                  </div>
                </div>
                <div className="chat-messages">
                  {messages.map(msg => (
                    <div key={msg.id} className={`message ${msg.type}`}>
                      <div className="message-avatar">
                        {msg.type === 'ai' ? '🤖' : '👤'}
                      </div>
                      <div className="message-content">
                        {msg.content.split('\n').map((line, i) => (
                          <div key={i}>{line.replace(/\*\*(.*?)\*\*/g, '$1')}</div>
                        ))}
                      </div>
                    </div>
                  ))}
                  {isTyping && (
                    <div className="message ai">
                      <div className="message-avatar">🤖</div>
                      <div className="message-content">
                        <div className="typing-indicator">
                          <span></span>
                          <span></span>
                          <span></span>
                        </div>
                      </div>
                    </div>
                  )}
                  <div ref={chatEndRef} />
                </div>
                <div className="chat-input-container">
                  <input
                    type="text"
                    className="chat-input"
                    placeholder="问我任何关于财务的问题..."
                    value={inputValue}
                    onChange={e => setInputValue(e.target.value)}
                    onKeyPress={e => e.key === 'Enter' && handleSendMessage()}
                  />
                  <button className="chat-send" onClick={handleSendMessage}>
                    <Send size={20} style={{ color: '#0a0e14' }} />
                  </button>
                </div>
              </div>
            </section>

            <section style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '1.5rem', marginBottom: '1.5rem' }}>
              <div className="card">
                <div className="card-header">
                  <div className="card-title">
                    <AlertTriangle size={20} style={{ color: '#fbbf24' }} />
                    风险预警
                  </div>
                </div>
                <div className="alert-list">
                  {alerts.map((alert, index) => (
                    <div key={index} className="alert-item">
                      <div className={`alert-icon ${alert.type}`}>
                        {alert.type === 'warning' ? '⚠️' : 'ℹ️'}
                      </div>
                      <div className="alert-content">
                        <h4>{alert.title}</h4>
                        <p>{alert.desc}</p>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              <div className="card">
                <div className="card-header">
                  <div className="card-title">
                    <Lightbulb size={20} style={{ color: '#d4af37' }} />
                    决策建议
                  </div>
                </div>
                <div className="suggestion-tags" style={{ marginBottom: '1rem', marginTop: 0 }}>
                  {suggestions.map((s, index) => (
                    <div key={index} style={{ marginBottom: '1rem' }}>
                      <div style={{ display: 'flex', alignItems: 'flex-start', gap: '1rem' }}>
                        <div style={{ fontSize: '1.5rem' }}>{s.icon}</div>
                        <div>
                          <div className="suggestion-title">{s.title}</div>
                          <div className="suggestion-desc">{s.desc}</div>
                          <div className="suggestion-tags">
                            {s.tags.map((tag, i) => (
                              <span key={i} className="suggestion-tag">{tag}</span>
                            ))}
                          </div>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </section>

            <section className="card">
              <div className="card-header">
                <div className="card-title">
                  <BarChart3 size={20} style={{ color: '#d4af37' }} />
                  利润趋势
                </div>
              </div>
              <div className="chart-container" style={{ height: '250px' }}>
                <ReactECharts option={getProfitBarOption()} style={{ height: '100%', width: '100%' }} />
              </div>
            </section>
          </>
        )}
      </main>
    </div>
  );
}

export default App;
