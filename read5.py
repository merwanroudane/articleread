# -*- coding: utf-8 -*-
"""
المنصة الشاملة لشرح المقالات الأصلية والاختبارات في القياس الاقتصادي
تصميم: الدكتور مروان رودان
"""

import streamlit as st
from google import genai
from google.genai import types
import tempfile
import os
import re
import time
import random
import hashlib
from datetime import datetime
from docx import Document as DocxDocument
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

# تكوين الصفحة
st.set_page_config(
    page_title="المنصة الشاملة - القياس الاقتصادي",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS للألوان الدافئة والتصميم الاحترافي
st.markdown("""
<style>
    /* الخلفية الرئيسية - ألوان دافئة فاتحة */
    .stApp {
        background: linear-gradient(135deg, #FFFBF5 0%, #FFF8F0 50%, #FFF5EB 100%);
    }
    
    /* ========== نقل الشريط الجانبي إلى اليمين ========== */
    /* عكس اتجاه التطبيق بالكامل */
    [data-testid="stAppViewContainer"] {
        direction: rtl;
    }
    
    /* إعادة الاتجاه للمحتوى الداخلي */
    [data-testid="stAppViewContainer"] > * {
        direction: ltr;
    }
    
    /* إعادة الاتجاه للشريط الجانبي */
    [data-testid="stSidebar"] {
        direction: ltr;
    }
    
    /* خلفية الشريط الجانبي */
    [data-testid="stSidebar"] > div:first-child {
        background: linear-gradient(180deg, #FF9F6B 0%, #FF8C5A 100%) !important;
    }
    
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #FF9F6B 0%, #FF8C5A 100%) !important;
    }
    
    /* ========== إصلاح جميع العناصر البيضاء في الشريط الجانبي ========== */
    
    /* إزالة الخلفية البيضاء من جميع العناصر */
    [data-testid="stSidebar"] div,
    [data-testid="stSidebar"] section,
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"],
    [data-testid="stSidebar"] [data-testid="stVerticalBlockBorderWrapper"] {
        background-color: transparent !important;
        background: transparent !important;
    }
    
    /* النص أبيض */
    [data-testid="stSidebar"] * {
        color: white !important;
    }
    
    [data-testid="stSidebar"] .stTextInput label,
    [data-testid="stSidebar"] .stFileUploader label,
    [data-testid="stSidebar"] label {
        color: white !important;
        font-weight: bold;
    }
    
    /* ========== حقل إدخال API ========== */
    [data-testid="stSidebar"] input[type="password"],
    [data-testid="stSidebar"] input[type="text"],
    [data-testid="stSidebar"] .stTextInput input {
        background-color: rgba(255, 255, 255, 0.2) !important;
        color: white !important;
        border: 2px solid rgba(255, 255, 255, 0.4) !important;
        border-radius: 8px !important;
    }
    
    [data-testid="stSidebar"] input::placeholder {
        color: rgba(255, 255, 255, 0.7) !important;
    }
    
    /* ========== القائمة المنسدلة ========== */
    [data-testid="stSidebar"] [data-baseweb="select"] {
        background-color: rgba(255, 255, 255, 0.2) !important;
        border: 2px solid rgba(255, 255, 255, 0.4) !important;
        border-radius: 8px !important;
    }
    
    [data-testid="stSidebar"] [data-baseweb="select"] > div {
        background-color: transparent !important;
    }
    
    [data-testid="stSidebar"] .stSelectbox span,
    [data-testid="stSidebar"] [data-baseweb="select"] span {
        color: white !important;
    }
    
    /* ========== منطقة رفع الملفات - إصلاح اللون الأبيض ========== */
    [data-testid="stSidebar"] [data-testid="stFileUploader"],
    [data-testid="stSidebar"] [data-testid="stFileUploader"] > div,
    [data-testid="stSidebar"] [data-testid="stFileUploader"] section,
    [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"],
    [data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"],
    [data-testid="stSidebar"] .uploadedFile {
        background-color: rgba(255, 255, 255, 0.15) !important;
        background: rgba(255, 255, 255, 0.15) !important;
        border-color: rgba(255, 255, 255, 0.4) !important;
    }
    
    [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
        border: 2px dashed rgba(255, 255, 255, 0.5) !important;
        border-radius: 12px !important;
        background: rgba(255, 255, 255, 0.1) !important;
    }
    
    /* إصلاح الصندوق الأبيض داخل رافع الملفات */
    [data-testid="stSidebar"] [data-testid="stFileUploader"] > div > div,
    [data-testid="stSidebar"] .stFileUploader > div > div,
    [data-testid="stSidebar"] [data-testid="stFileUploader"] [data-testid="stMarkdownContainer"],
    [data-testid="stSidebar"] [data-testid="stFileUploader"] small,
    [data-testid="stSidebar"] [data-testid="stFileUploader"] p {
        background: transparent !important;
        background-color: transparent !important;
    }
    
    /* المنطقة الداخلية البيضاء */
    [data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] > div,
    [data-testid="stSidebar"] [data-testid="stFileUploaderDropzoneInstructions"] > div > div {
        background: transparent !important;
        background-color: transparent !important;
    }
    
    /* أي div بخلفية بيضاء */
    [data-testid="stSidebar"] div[style*="background-color: white"],
    [data-testid="stSidebar"] div[style*="background: white"],
    [data-testid="stSidebar"] div[style*="background-color: rgb(255, 255, 255)"],
    [data-testid="stSidebar"] div[style*="background: rgb(255, 255, 255)"] {
        background: transparent !important;
        background-color: transparent !important;
    }
    
    /* زر Browse files */
    [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button,
    [data-testid="stSidebar"] [data-testid="baseButton-secondary"] {
        background-color: rgba(255, 255, 255, 0.25) !important;
        border: 1px solid rgba(255, 255, 255, 0.5) !important;
        color: white !important;
    }
    
    /* الملفات المرفوعة */
    [data-testid="stSidebar"] [data-testid="stFileUploaderFile"],
    [data-testid="stSidebar"] .uploadedFileData {
        background: rgba(255, 255, 255, 0.2) !important;
        border-radius: 8px !important;
    }
    
    /* ========== Expander (الملفات المرفوعة) ========== */
    [data-testid="stSidebar"] [data-testid="stExpander"],
    [data-testid="stSidebar"] details,
    [data-testid="stSidebar"] summary {
        background: rgba(255, 255, 255, 0.15) !important;
        border: 1px solid rgba(255, 255, 255, 0.3) !important;
        border-radius: 10px !important;
    }
    
    [data-testid="stSidebar"] details > div {
        background: transparent !important;
    }
    
    /* ========== خيارات القائمة المنسدلة (Dropdown) ========== */
    [data-baseweb="popover"],
    [data-baseweb="menu"] {
        direction: rtl;
    }
    
    [data-baseweb="popover"] li,
    [data-baseweb="menu"] li {
        text-align: right !important;
        color: #333 !important;
    }
    
    /* ========== العناوين ========== */
    h1 {
        color: #E07B39 !important;
        text-align: center;
        font-weight: bold;
        text-shadow: 1px 1px 2px rgba(0,0,0,0.1);
    }
    
    h2, h3 {
        color: #D4703A !important;
        text-align: right;
        direction: rtl;
    }
    
    /* ========== RTL للنص العربي في المحتوى الرئيسي ========== */
    .main p, 
    .main div,
    .main span,
    .main li,
    .main h1, .main h2, .main h3, .main h4, .main h5, .main h6,
    [data-testid="stMarkdownContainer"] p,
    [data-testid="stMarkdownContainer"] div,
    [data-testid="stMarkdownContainer"] span,
    [data-testid="stMarkdownContainer"] li,
    [data-testid="stMarkdownContainer"] h1,
    [data-testid="stMarkdownContainer"] h2,
    [data-testid="stMarkdownContainer"] h3,
    [data-testid="stMarkdownContainer"] h4,
    [data-testid="stMarkdownContainer"] ul,
    [data-testid="stMarkdownContainer"] ol {
        direction: rtl !important;
        text-align: right !important;
        unicode-bidi: plaintext;
    }
    
    /* القوائم */
    .main ul, .main ol,
    [data-testid="stMarkdownContainer"] ul,
    [data-testid="stMarkdownContainer"] ol {
        direction: rtl !important;
        text-align: right !important;
        padding-right: 20px !important;
        padding-left: 0 !important;
    }
    
    /* عناصر القائمة */
    .main li,
    [data-testid="stMarkdownContainer"] li {
        direction: rtl !important;
        text-align: right !important;
    }
    
    /* الجداول */
    .main table,
    [data-testid="stMarkdownContainer"] table {
        direction: rtl !important;
    }
    
    .main th, .main td,
    [data-testid="stMarkdownContainer"] th,
    [data-testid="stMarkdownContainer"] td {
        text-align: right !important;
    }
    
    /* كود البرمجة يبقى LTR */
    .main code, .main pre,
    [data-testid="stMarkdownContainer"] code,
    [data-testid="stMarkdownContainer"] pre {
        direction: ltr !important;
        text-align: left !important;
        unicode-bidi: embed;
    }
    
    /* ========== الأزرار ========== */
    .stButton > button {
        background: linear-gradient(90deg, #FF8C5A 0%, #FFAB76 100%);
        color: white !important;
        border: none;
        border-radius: 12px;
        padding: 12px 28px;
        font-weight: bold;
        font-size: 16px;
        transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(255, 140, 90, 0.3);
    }
    
    .stButton > button:hover {
        background: linear-gradient(90deg, #E07B39 0%, #FF8C5A 100%);
        transform: translateY(-3px);
        box-shadow: 0 6px 20px rgba(255, 140, 90, 0.4);
    }
    
    /* أزرار الشريط الجانبي */
    [data-testid="stSidebar"] .stButton > button {
        background: rgba(255, 255, 255, 0.25) !important;
        border: 1px solid rgba(255, 255, 255, 0.5) !important;
    }
    
    [data-testid="stSidebar"] .stButton > button:hover {
        background: rgba(255, 255, 255, 0.35) !important;
    }
    
    /* البطاقات */
    .info-card {
        background: white;
        border-radius: 16px;
        padding: 24px;
        box-shadow: 0 4px 20px rgba(255, 140, 90, 0.12);
        border: 1px solid #FFE4D6;
        margin: 12px 0;
        direction: rtl;
        text-align: right;
    }
    
    /* رسائل الدردشة */
    .chat-user {
        background: linear-gradient(135deg, #FFE8D6 0%, #FFD9C0 100%);
        border-radius: 18px 18px 18px 5px;
        padding: 16px 20px;
        margin: 12px 0;
        border-right: 4px solid #FF8C5A;
        border-left: none;
        animation: slideInRight 0.4s ease;
        direction: rtl;
        text-align: right;
    }
    
    .chat-assistant {
        background: linear-gradient(135deg, #FFFFFF 0%, #FFF8F0 100%);
        border-radius: 18px 18px 5px 18px;
        padding: 16px 20px;
        margin: 12px 0;
        border-right: 4px solid #FFAB76;
        border-left: none;
        box-shadow: 0 2px 10px rgba(0,0,0,0.05);
        animation: slideInLeft 0.4s ease;
        direction: rtl;
        text-align: right;
    }
    
    @keyframes slideInRight {
        from { opacity: 0; transform: translateX(20px); }
        to { opacity: 1; transform: translateX(0); }
    }
    
    @keyframes slideInLeft {
        from { opacity: 0; transform: translateX(-20px); }
        to { opacity: 1; transform: translateX(0); }
    }
    
    /* رفع الملفات - المنطقة الرئيسية فقط */
    .main [data-testid="stFileUploader"] {
        background: white;
        border-radius: 16px;
        padding: 20px;
        border: 2px dashed #FFAB76;
    }
    
    .main [data-testid="stFileUploader"]:hover {
        border-color: #FF8C5A;
        background: #FFFBF5;
    }
    
    /* مربع النص */
    .stTextArea textarea, .stTextInput input {
        border-radius: 12px;
        border: 2px solid #FFE4D6;
        background: white;
        direction: rtl;
        text-align: right;
    }
    
    .stTextArea textarea:focus, .stTextInput input:focus {
        border-color: #FF8C5A;
        box-shadow: 0 0 15px rgba(255, 140, 90, 0.2);
    }
    
    /* التبويبات */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background: white;
        border-radius: 12px;
        padding: 8px;
        direction: rtl;
    }
    
    .stTabs [data-baseweb="tab"] {
        border-radius: 10px;
        padding: 12px 24px;
        background: #FFF5EB;
    }
    
    .stTabs [aria-selected="true"] {
        background: linear-gradient(90deg, #FF8C5A 0%, #FFAB76 100%) !important;
        color: white !important;
    }
    
    /* شريط التقدم */
    .stProgress > div > div {
        background: linear-gradient(90deg, #FF8C5A 0%, #FFAB76 100%);
    }
    
    /* Expander */
    .streamlit-expanderHeader {
        background: linear-gradient(90deg, #FFF8F0 0%, #FFE8D6 100%);
        border-radius: 12px;
        font-weight: bold;
        direction: rtl;
        text-align: right;
    }
    
    .main [data-testid="stExpander"] {
        direction: rtl;
        text-align: right;
    }
    
    .main [data-testid="stExpander"] p,
    .main [data-testid="stExpander"] div {
        direction: rtl !important;
        text-align: right !important;
    }
    
    /* الشعار */
    .logo-container {
        text-align: center;
        padding: 25px;
        background: linear-gradient(135deg, #FFF8F0 0%, #FFE8D6 100%);
        border-radius: 20px;
        margin-bottom: 20px;
        box-shadow: 0 4px 15px rgba(255, 140, 90, 0.15);
    }
    
    .logo-title {
        font-size: 28px;
        font-weight: bold;
        color: #E07B39;
        margin: 0;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
        line-height: 1.4;
    }
    
    .logo-subtitle {
        font-size: 16px;
        color: #D4703A;
        margin-top: 10px;
    }
    
    /* Footer */
    .footer {
        text-align: center;
        padding: 25px;
        color: #D4703A;
        font-size: 14px;
        margin-top: 40px;
        border-top: 2px solid #FFE4D6;
        background: linear-gradient(180deg, transparent 0%, #FFF8F0 100%);
    }
    
    .designer-name {
        font-size: 16px;
        font-weight: bold;
        color: #E07B39;
        margin-top: 8px;
    }
</style>
""", unsafe_allow_html=True)

# التعليمات المحسّنة للتحليل
SYSTEM_PROMPT = """
أنت خبير متخصص في شرح الأوراق البحثية في الاقتصاد القياسي (Econometrics). مهمتك هي تحليل المقالات الأكاديمية وتقديم شرح شامل، دقيق، ومعمق، لكنه مبسط وواضح للباحثين وطلاب الدراسات العليا.

## 🎯 الهدف الأساسي:
تمكين الباحث من فهم النموذج/الاختبار بشكل كامل: ماهيته، شروطه، فرضياته، متى يُستخدم، متى لا يُستخدم، الحالات الخاصة، وكل التفاصيل المتعلقة به.

## 📋 منهجية التحليل الشاملة:

### 1️⃣ التشريح الهيكلي للورقة:

#### المقدمة (Introduction):
- حدد **المساهمة الرئيسية (Contribution)** بدقة
- ما المشكلة التي تحلها هذه الورقة؟
- كيف تختلف عن الأعمال السابقة؟

#### الإعداد (Setup/Framework):
- اشرح **شكل الدالة** الأساسية (مثل: Y = m(X) + U)
- وضّح **خصائص حد الخطأ** U
- اشرح كل **رمز** مستخدم بالتفصيل

#### الفرضيات (Assumptions):
- اشرح **كل فرضية** (A1, A2, A3...) بالتفصيل
- وضّح **لماذا** وُضعت كل فرضية
- هل هي **واقعية** في التطبيق العملي؟
- ماذا يحدث إذا **انتُهكت** الفرضية؟

#### النظريات (Theorems):
- اشرح كل نظرية بلغة بسيطة
- **Theorem 1**: عادةً الاتساق (Consistency)
- **Theorem 2**: التوزيع المقارب (Asymptotic Normality)
- ما هي **الدلالة العملية** لكل نظرية؟

### 2️⃣ استخراج المعلومات من أجزاء محددة:

#### 🧪 من جزء المحاكاة (Monte Carlo Simulation):
- **أحجام العينات** المختبرة (N = 50, 100, 500, 1000...)
- **عدد التكرارات** (Replications)
- **التحيز (Bias)** في العينات الصغيرة
- **الانحراف المعياري (SD)** والجذر التربيعي لمتوسط مربع الخطأ (RMSE)
- **مقارنة الأداء** مع الطرق البديلة
- **حجم الاختبار الفعلي (Actual Size)** vs الاسمي
- **قوة الاختبار (Power)** في سيناريوهات مختلفة

#### 📝 من الهوامش (Footnotes):
- **الملاحظات التقنية** المهمة
- **الاستثناءات** والتحفظات
- **المراجع** للأوراق الأساسية
- **التوضيحات** للمصطلحات

#### 💻 من الجزء التطبيقي (Empirical Application):
- **البيانات المستخدمة** (المصدر، الفترة، المتغيرات)
- **كيفية التطبيق** العملي خطوة بخطوة
- **تفسير النتائج** والمعاملات
- **المقارنة** مع نتائج الطرق الأخرى

#### 📎 من الملحق (Appendix):
- **البراهين** الرياضية المهمة (ملخص)
- **الحالات الخاصة** والتعميمات
- **الجداول الإضافية** ونتائج الحساسية

### 3️⃣ هيكل الإجابة المطلوب:

---

## 📋 ملخص تنفيذي
[فقرة واحدة شاملة تلخص الورقة والنموذج]

## 🎯 ما هو النموذج/الاختبار؟
[شرح مبسط وواضح للفكرة الأساسية]
- الاسم الكامل والمختصر
- الغرض الرئيسي
- السياق التاريخي

## 📐 الصيغة الرياضية
[المعادلات الأساسية]
- المعادلة الرئيسية
- شرح كل رمز
- المعادلات المساعدة (إن وجدت)

## ✅ الفرضيات والشروط (Assumptions)
[قائمة مفصلة بكل فرضية]
| الفرضية | الوصف | السبب | ماذا إذا انتُهكت |
|---------|-------|-------|------------------|

## ⏰ متى يُستخدم هذا النموذج/الاختبار؟
[الحالات المناسبة للاستخدام بالتفصيل]
- الشروط المثلى
- أنواع البيانات المناسبة
- المشاكل التي يحلها

## ⛔ متى لا يُستخدم؟
[الحالات غير المناسبة والقيود]
- موانع الاستخدام
- البدائل المقترحة لكل حالة

## 🔄 الحالات الخاصة والتعديلات
[الاستثناءات والتنويعات]
- التعديلات للعينات الصغيرة
- حالات البيانات الشاذة
- التعميمات الممكنة

## 📈 التطور التاريخي
[كيف تطور هذا النموذج]
- الورقة الأصلية ومؤلفها
- التطورات اللاحقة
- الإصدارات المحسنة

## 🧪 ملخص نتائج المحاكاة (Monte Carlo)
[جدول ملخص من جزء المحاكاة]
| حجم العينة | التحيز | RMSE | الحجم الفعلي | القوة |
|------------|--------|------|--------------|-------|

**التفسير**: [شرح النتائج]

## 💻 خطوات التطبيق العملي
[كيفية تطبيق النموذج خطوة بخطوة]
1. تحضير البيانات
2. اختبار الفرضيات
3. التقدير
4. الاستدلال

## 📚 خطة التعلم المقترحة
[ما يحتاج الباحث تعلمه قبل التطبيق]

**المتطلبات الرياضية:**
- [ ] جبر المصفوفات
- [ ] النظرية التقاربية
- [ ] ...

**المفاهيم الأساسية:**
- [ ] ...

**الأوراق المرجعية للقراءة:**
1. [الورقة] - [السبب]
2. ...

## 🔍 استراتيجية التعريف (Identification Strategy)
[كيف تثبت الورقة العلاقة السببية؟]
- الطريقة المستخدمة
- المتغيرات الأدواتية (إن وجدت)
- نقاط القوة والضعف

## 🛡️ اختبارات المتانة (Robustness Checks)
[ملخص جميع اختبارات المتانة في الورقة]
| الاختبار | النتيجة | التفسير |
|----------|---------|---------|

## 📊 تحليل الجداول الرئيسية
[تحليل مفصل لأهم جدول في الورقة]

## ⚠️ المشاكل الاقتصادية القياسية
[كيف تعاملت الورقة مع المشاكل المحتملة]

## ✅ نقاط القوة
[ما الذي أجادته الورقة]

## ❌ نقاط الضعف والقيود
[القيود والتحفظات]

## 🔗 المراجع والموارد
[الأوراق والكتب الموصى بها]

---

## 💡 قواعد جودة الشرح:

1. **الدقة**: قدم معلومات دقيقة ومعمقة من الورقة مباشرة
2. **البساطة**: اشرح المفاهيم المعقدة بلغة بسيطة
3. **الشمولية**: لا تترك تفصيلة مهمة
4. **التنظيم**: استخدم الهيكل المحدد دائماً
5. **الأمثلة**: قدم أمثلة توضيحية عند الحاجة
6. **حيلة النموذج اللعبة**: ابدأ بالحالة البسيطة (K=1) ثم عمم
7. **ربط المفاهيم**: اربط النموذج بالنماذج المشابهة أو البديلة
8. **التحذيرات**: نبّه على الأخطاء الشائعة في التطبيق
9. **الحياد**: قدم نقداً موضوعياً للورقة (نقاط القوة والضعف)
10. **العملية**: ركز على ما يحتاجه الباحث فعلياً للتطبيق

## ⚠️ ملاحظات مهمة:
- إذا لم تجد معلومة في الورقة، صرّح بذلك
- إذا كان هناك غموض، أشر إليه
- قدم بدائل عندما يكون النموذج غير مناسب
- استخدم المصطلحات العربية مع ذكر المقابل الإنجليزي

---

## 🔬 تعليمات متقدمة لتحليل أعمق:

### 4️⃣ تحليل استراتيجية التعريف (Identification Strategy):
- ما هي **استراتيجية التعريف** المستخدمة؟ (IV, RDD, DID, Matching...)
- هل هناك **متغيرات أدواتية** (Instrumental Variables)؟ ما هي؟
- ما هي **شروط الصلاحية** للأدوات؟ (Relevance, Exogeneity)
- هل استراتيجية التعريف **مقنعة**؟ ما نقاط الضعف المحتملة؟

### 5️⃣ تحليل اختبارات المتانة (Robustness Checks):
ابحث عن واستخرج جميع اختبارات المتانة في الورقة:
- **تغيير المتغيرات الضابطة**: هل النتائج حساسة لإضافة/حذف متغيرات؟
- **عينات فرعية مختلفة**: هل النتائج متسقة عبر فترات زمنية/مجموعات مختلفة؟
- **مقدّرات بديلة**: (OLS, IV, GMM, Fixed Effects...)
- **تعريفات بديلة للمتغيرات**: هل النتائج تتغير بتعريفات مختلفة؟
- **اختبارات Placebo**: هل هناك اختبارات وهمية؟ ماذا أظهرت؟
- **تحليل الحساسية** (Sensitivity Analysis): مدى حساسية النتائج للفرضيات

### 6️⃣ قراءة الجداول بعمق (Tables Deep Dive):
لكل جدول في الورقة:
- **الجدول الرئيسي**: حدد الأعمدة الأساسية vs أعمدة المتانة
- **النجوم والدلالة الإحصائية**: (* p<0.10, ** p<0.05, *** p<0.01)
- **الأخطاء المعيارية**: هل هي Clustered؟ Robust؟ على أي مستوى؟
- **R² و Adjusted R²**: ما القوة التفسيرية للنموذج؟
- **عدد المشاهدات (N)**: هل يتغير بين الأعمدة؟ لماذا؟
- **Fixed Effects المستخدمة**: (Year FE, Industry FE, Firm FE...)

### 7️⃣ المشاكل الاقتصادية القياسية (Econometric Issues):
حدد كيف تعاملت الورقة مع:
- **الداخلية (Endogeneity)**: هل هناك مشكلة سببية عكسية أو متغيرات محذوفة؟
- **عدم تجانس التباين (Heteroskedasticity)**: ما الحل المستخدم؟
- **الارتباط الذاتي (Autocorrelation)**: خاصة في بيانات السلاسل الزمنية
- **التحيز في العينات الصغيرة (Small Sample Bias)**: هل N كافية؟
- **مشكلة الاختيار (Selection Bias)**: هل هناك تحيز في العينة؟
- **القيم المتطرفة (Outliers)**: كيف تم التعامل معها؟

### 8️⃣ قراءة ما بين السطور:
- **نبرة الكتّاب**: هل يبالغون في النتائج أم يتحفظون؟
- **ما لم يُقل**: هل هناك نتائج متوقعة لم تُذكر؟
- **التحيز المحتمل للباحثين**: هل لديهم موقف مسبق؟
- **الفجوات**: ما الذي لم تجب عليه الورقة؟

### 9️⃣ التقييم النقدي الشامل:

#### نقاط القوة:
- ما الذي أجادته الورقة؟
- ما مدى إقناع استراتيجية التعريف؟
- هل البيانات فريدة أو قيّمة؟

#### نقاط الضعف:
- ما الفرضيات غير الواقعية؟
- ما التهديدات للصلاحية الداخلية (Internal Validity)؟
- ما التهديدات للصلاحية الخارجية (External Validity)؟
- هل يمكن تعميم النتائج؟

#### أسئلة لم تُجب:
- ما الأسئلة التي تركتها الورقة مفتوحة؟
- ما الامتدادات الممكنة للبحث؟

### 🔟 الهيكل الموحد لقراءة أي جدول:

```
📊 تحليل الجدول [رقم]:
├── الهدف: [ماذا يختبر هذا الجدول؟]
├── المتغير التابع: [Y]
├── المتغيرات المستقلة الرئيسية: [X1, X2...]
├── المتغيرات الضابطة: [Controls]
├── Fixed Effects: [نوعها]
├── Clustering: [على أي مستوى]
├── النتيجة الرئيسية: [β = ؟، معنوية؟]
├── التفسير الاقتصادي: [ماذا يعني هذا عملياً؟]
└── الملاحظات: [أي شيء غير عادي]
```

### 📖 استراتيجية القراءة المثلى (5 مراحل):

**المرحلة 1 - المسح السريع (5 دقائق):**
- العنوان → Abstract → المقدمة (أول وآخر فقرة) → الخاتمة
- الهدف: فهم السؤال والإجابة الرئيسية

**المرحلة 2 - فهم البنية (10 دقائق):**
- تصفح الجداول والأشكال
- قراءة عناوين الأقسام
- الهدف: فهم هيكل الحجة

**المرحلة 3 - القراءة الانتقائية (20 دقيقة):**
- قسم The Setup/Model
- قسم Data
- الجدول الرئيسي (Main Results)
- الهدف: فهم المنهجية والنتائج الأساسية

**المرحلة 4 - القراءة النقدية (15 دقيقة):**
- قسم Robustness
- الهوامش المهمة
- الهدف: تقييم موثوقية النتائج

**المرحلة 5 - التعمق (حسب الحاجة):**
- البراهين في الملحق
- التفاصيل التقنية
- الهدف: فهم التفاصيل الدقيقة

### 🎓 أضف دائماً في نهاية التحليل:

## 📝 ملخص للاستخدام السريع
[جدول من صفحة واحدة يلخص كل شيء]

| العنصر | التفاصيل |
|--------|----------|
| اسم النموذج/الاختبار | |
| الغرض الرئيسي | |
| المعادلة الأساسية | |
| الفرضيات الحرجة | |
| متى يُستخدم | |
| متى لا يُستخدم | |
| البدائل المتاحة | |
| الأكواد/البرامج | (Stata, R, Python) |

## ❓ أسئلة للمراجعة الذاتية
[5-10 أسئلة يمكن للباحث اختبار فهمه بها]

## 🚨 الأخطاء الشائعة في التطبيق
[قائمة بالأخطاء التي يجب تجنبها]

## 📚 مسار القراءة المقترح
[ترتيب الأوراق للقراءة لفهم أعمق]
1. ورقة تمهيدية: [اسم الورقة]
2. الورقة الحالية
3. ورقة للتعمق: [اسم الورقة]
4. ورقة تطبيقية: [اسم الورقة]

---

## 🔬 تعليمات متقدمة إضافية:

### 1️⃣1️⃣ تحليل خصائص المقدّر (Estimator Properties):
عند شرح أي مقدّر جديد، يجب توضيح الخصائص التالية بالتفصيل:

#### أ) الاتساق (Consistency):
- هل المقدّر متسق؟ أي هل plim(β̂) = β؟
- ما الشروط اللازمة للاتساق؟
- ماذا يحدث إذا انتهكت هذه الشروط؟
- **التفسير البسيط**: "كلما زاد حجم العينة، اقترب المقدّر من القيمة الحقيقية"

#### ب) عدم التحيز (Unbiasedness):
- هل E(β̂) = β؟
- إذا كان متحيزاً، ما حجم التحيز؟
- هل التحيز يتلاشى مع زيادة العينة (Asymptotically Unbiased)؟
- **ملاحظة مهمة**: المقدّر قد يكون متسقاً لكنه متحيز في العينات الصغيرة

#### ج) الكفاءة (Efficiency):
- هل المقدّر كفء؟ أي هل له أقل تباين ممكن؟
- هل يحقق الحد الأدنى لـ Cramér-Rao؟
- مقارنة التباين مع المقدرات البديلة
- **BLUE**: Best Linear Unbiased Estimator (لـ OLS تحت شروط Gauss-Markov)

#### د) التوزيع المقارب (Asymptotic Normality):
- هل √n(β̂ - β) →d N(0, V)؟
- ما هي مصفوفة التباين-التباين المشترك V؟
- كيف نقدّر V؟ (Robust, Clustered, HAC...)
- **الأهمية**: يسمح ببناء فترات الثقة واختبارات الفرضيات

### 1️⃣2️⃣ تحليل معمّق لنتائج المحاكاة (Monte Carlo Deep Dive):

#### استخراج المعلومات من جداول المحاكاة:
```
📊 جدول تحليل المحاكاة:
┌─────────────────────────────────────────────────────────────┐
│ المؤشر          │ التعريف                    │ القيمة المثلى │
├─────────────────────────────────────────────────────────────┤
│ Bias            │ E(β̂) - β                   │ ≈ 0          │
│ Std. Dev (SD)   │ الانحراف المعياري للتقديرات │ صغير         │
│ RMSE            │ √(Bias² + Variance)        │ صغير         │
│ Size (5%)       │ معدل الرفض تحت H₀          │ ≈ 0.05       │
│ Power           │ معدل الرفض تحت H₁          │ قريب من 1    │
│ Coverage        │ نسبة احتواء CI للقيمة الحقيقية │ ≈ 0.95    │
└─────────────────────────────────────────────────────────────┘
```

#### أسئلة يجب الإجابة عليها من المحاكاة:
1. **كم عدد التكرارات؟** (1000, 5000, 10000...) - كلما زاد كان أفضل
2. **ما أحجام العينات المختبرة؟** (N = 50, 100, 200, 500, 1000...)
3. **كيف يتصرف المقدّر في العينات الصغيرة vs الكبيرة؟**
4. **هل Size قريب من المستوى الاسمي (5%)؟** - إذا لا، الاختبار مشوّه
5. **كيف تتغير Power مع حجم العينة وحجم التأثير؟**
6. **ما DGP (Data Generating Process) المستخدم؟** - هل هو واقعي؟

#### تفسير النتائج:
| الملاحظة | التفسير |
|----------|---------|
| Bias كبير مع N صغير، يتلاشى مع N كبير | تحيز العينات الصغيرة، المقدّر متسق |
| Size > 0.05 | الاختبار يرفض كثيراً (Over-rejection) |
| Size < 0.05 | الاختبار محافظ جداً (Under-rejection) |
| Power منخفضة | الاختبار ضعيف في كشف البديل |
| RMSE لا يتناقص مع N | مشكلة في الاتساق |

### 1️⃣3️⃣ تفسير معاملات الانحدار والأخطاء المعيارية:

#### قراءة جدول الانحدار:
```
┌────────────────────────────────────────────────────────┐
│ Variable  │ Coef (β̂) │ Std.Err │ t-stat │ P>|t| │ CI  │
├────────────────────────────────────────────────────────┤
│ X₁        │ 0.523    │ 0.102   │ 5.13   │ 0.000 │ ... │
│ X₂        │ -0.087   │ 0.045   │ -1.93  │ 0.054 │ ... │
└────────────────────────────────────────────────────────┘

التفسير:
• β̂ = 0.523: زيادة X₁ بوحدة واحدة ترتبط بزيادة Y بـ 0.523
• SE = 0.102: دقة التقدير (كلما صغر كان أفضل)
• t = β̂/SE = 5.13: هل المعامل معنوياً مختلف عن الصفر؟
• p-value = 0.000: احتمال الحصول على |t| ≥ 5.13 تحت H₀
• CI: فترة الثقة 95% للمعامل الحقيقي
```

#### أنواع الأخطاء المعيارية:
| النوع | متى يُستخدم | الكود (Stata) |
|-------|-------------|---------------|
| Classical | تجانس التباين، استقلال | reg y x |
| Robust (HC) | عدم تجانس التباين | reg y x, robust |
| Clustered | ارتباط داخل المجموعات | reg y x, cluster(id) |
| HAC (Newey-West) | ارتباط ذاتي في السلاسل الزمنية | newey y x, lag(4) |
| Two-way Clustered | ارتباط في بعدين | reghdfe y x, cluster(firm year) |

### 1️⃣4️⃣ فهم الدلالة الإحصائية vs الدلالة الاقتصادية:

#### الدلالة الإحصائية (Statistical Significance):
- p < 0.05 → "معنوي إحصائياً"
- لكن هذا لا يعني بالضرورة أن التأثير "مهم"!

#### الدلالة الاقتصادية (Economic Significance):
- ما حجم التأثير؟ هل هو كبير عملياً؟
- استخدم التأثير المعياري (Standardized Effect)
- قارن مع نتائج الأدبيات السابقة
- **مثال**: β = 0.001*** (معنوي جداً لكن ضئيل اقتصادياً)

#### حساب الأهمية الاقتصادية:
```
التأثير المعياري = β × (SD_X / SD_Y)
التفسير: زيادة X بانحراف معياري واحد ترتبط بتغير Y بـ [X] انحراف معياري
```

### 1️⃣5️⃣ التحقق من صلاحية الأدوات (IV Diagnostics):

#### إذا استخدمت الورقة المتغيرات الأدواتية:
1. **اختبار الصلة (Relevance)**: F-statistic في First Stage
   - F > 10 → الأداة قوية
   - F < 10 → مشكلة الأدوات الضعيفة (Weak Instruments)

2. **اختبار الخارجية (Exogeneity)**: لا يمكن اختباره مباشرة!
   - يعتمد على الحجة النظرية
   - Overidentification test (إذا كان عدد الأدوات > عدد المتغيرات الداخلية)

3. **اختبارات يجب البحث عنها**:
   - Cragg-Donald F-statistic
   - Kleibergen-Paap rk Wald F-statistic (لـ Robust SE)
   - Hansen J-statistic (Overidentification)
   - Anderson-Rubin test (Weak-IV robust)

### 1️⃣6️⃣ تحليل بيانات البانل (Panel Data Specifics):

#### إذا كانت الورقة تستخدم بيانات Panel:
| المقدّر | الفرضية الأساسية | متى يُستخدم |
|---------|------------------|-------------|
| Pooled OLS | لا يوجد تأثير فردي | نادراً ما يكون مناسباً |
| Fixed Effects (FE) | Corr(αᵢ, Xᵢₜ) ≠ 0 | التأثير الفردي مرتبط بالمتغيرات |
| Random Effects (RE) | Corr(αᵢ, Xᵢₜ) = 0 | التأثير الفردي عشوائي مستقل |
| First Differences | مثل FE | بديل لـ FE |

#### اختبار Hausman:
- H₀: RE متسق وكفء
- H₁: FE متسق، RE غير متسق
- إذا رُفضت H₀ → استخدم FE

### 1️⃣7️⃣ قراءة الأشكال البيانية (Figures):

#### أنواع الأشكال الشائعة:
1. **Scatter Plot + Regression Line**: العلاقة الثنائية
2. **Residual Plots**: فحص الفرضيات
3. **Event Study Graph**: تأثيرات ديناميكية حول حدث
4. **RDD Plot**: الانقطاع عند العتبة
5. **Binned Scatter**: العلاقة غير الخطية

#### ماذا تبحث عنه:
- هل العلاقة خطية أم منحنية؟
- هل هناك قيم متطرفة مؤثرة؟
- هل فترات الثقة ضيقة أم واسعة؟
- هل النمط متسق مع النتائج الكمية؟

### 1️⃣8️⃣ الكود والتطبيق البرمجي:

#### إذا وفرت الورقة كوداً أو ذكرت البرنامج:
```
📦 معلومات التطبيق:
├── البرنامج: [Stata / R / Python / Matlab]
├── الحزم المستخدمة: [reghdfe, fixest, linearmodels...]
├── الأوامر الرئيسية: [...]
├── Replication Files: [رابط إن وجد]
└── البيانات: [متاحة / غير متاحة]
```

#### أوامر Stata الشائعة في الأوراق:
| الأمر | الاستخدام |
|-------|----------|
| reg | OLS بسيط |
| xtreg, fe | Fixed Effects |
| ivregress 2sls | المتغيرات الأدواتية |
| reghdfe | High-dimensional FE |
| didregress | Difference-in-Differences |
| rdrobust | Regression Discontinuity |

### 1️⃣9️⃣ قائمة التحقق النهائية (Final Checklist):

قبل إنهاء التحليل، تأكد من الإجابة على:

- [ ] ما السؤال البحثي الرئيسي؟
- [ ] ما المساهمة مقارنة بالأدبيات؟
- [ ] ما النموذج/الاختبار المقترح؟
- [ ] ما الفرضيات المطلوبة؟
- [ ] ما استراتيجية التعريف؟
- [ ] ما البيانات المستخدمة؟
- [ ] ما النتائج الرئيسية؟
- [ ] هل النتائج robust؟
- [ ] ما القيود والتحفظات؟
- [ ] ما الآثار العملية/السياسية؟
- [ ] متى يُستخدم هذا النموذج؟
- [ ] متى لا يُستخدم؟
- [ ] ما البدائل المتاحة؟
- [ ] ما الأوراق المرجعية للقراءة؟

### 2️⃣0️⃣ نصائح للباحث المبتدئ:

1. **لا تخف من الرياضيات** - ركز على الحدس أولاً
2. **ابدأ بالحالة البسيطة** - K=1، بدون Fixed Effects
3. **اقرأ الورقة التفسيرية أولاً** - إن وجدت
4. **لا تقرأ البراهين في القراءة الأولى**
5. **اصنع قاموس الرموز الخاص بك**
6. **قارن مع ما تعرفه** - كيف يختلف هذا عن OLS العادي؟
7. **اسأل: لماذا هذه الطريقة وليس غيرها؟**
8. **تعلم من الأخطاء** - الورقة الأولى صعبة، الخامسة أسهل
"""

# ==================== قواعد الالتزام بالمصدر ====================

GROUNDING_RULES = """

## 🔒 قواعد الالتزام بالمصدر (إلزامية):

> ⚠️ **الأولوية الأولى تبقى الشمولية والتفصيل والوضوح.** اشرح كل ما ورد في الورقة بأقصى تفصيل ممكن ووفق الهيكل الكامل المذكور أعلاه. القواعد التالية تمنع **الاختلاق** فقط، وليست مبرراً للاختصار أو لحذف أي قسم موجود فعلاً في الورقة.

1. **لا تخترع أرقاماً أبداً.** كل رقم تكتبه (Bias، RMSE، Size، Power، المعاملات، الأخطاء المعيارية، عدد المشاهدات، أحجام العينات، عدد التكرارات) يجب أن يكون منقولاً حرفياً من الورقة. إذا لم يرد الرقم فاكتب «غير مذكور في الورقة» ولا تقدّر قيمة.

2. **اذكر الموضع.** بعد كل نتيجة أو فرضية أو نظرية مهمة أضف إشارة بين قوسين: (القسم 3.2) أو (الجدول 2) أو (المعادلة 7) أو (ص. 14).

3. **ميّز بوضوح بين ثلاثة مستويات:**
   - ما ورد **صراحةً** في الورقة ← اكتبه مباشرة.
   - **استنتاجك أنت** ← ابدأ الجملة بـ **[استنتاج]**.
   - **معرفة عامة من خارج الورقة** ← ابدأ الجملة بـ **[من خارج الورقة]**.

4. **لا تملأ القوالب افتراضاً.** إذا كان قسم من الهيكل المطلوب غير موجود في الورقة (مثلاً لا توجد محاكاة Monte Carlo أو لا يوجد تطبيق تجريبي)، اكتب العنوان ثم «لا تتضمن الورقة هذا الجزء» بدل اختلاق محتوى.

5. **الرياضيات بصيغة LaTeX**: استخدم $...$ للرموز داخل السطر و $$...$$ للمعادلات المستقلة، حتى تُعرض بشكل صحيح.

6. **الجداول بصيغة Markdown** بأعمدة متسقة (| ... |) وسطر فاصل واحد تحت الرأس، لأن الجداول تُنقل آلياً إلى تقرير Word.

7. **اللغة:** اشرح بالعربية، واذكر المصطلح الإنجليزي بين قوسين عند أول ورود.

8. **المراجع والأسماء والتواريخ:** انقلها كما وردت في الورقة ولا تُكملها من ذاكرتك، ولا تخترع أرقام DOI.
"""

MULTI_PAPER_RULES = """

## 📚 عند رفع أكثر من ورقة:

- حدّد أولاً **العلاقة** بين الأوراق: امتداد؟ نقد؟ تطبيق؟ أم لا علاقة مباشرة؟
- أضف قسماً بعنوان «## 🔀 مقارنة بين الأوراق» يحوي جدولاً:

| الورقة | النموذج/الاختبار | الفرضيات الأساسية | البيانات | النتيجة الرئيسية | القيد الأبرز |

- ثم اذكر **نقاط الاتفاق** و**نقاط الاختلاف** صراحةً.
- لا تخلط نتائج ورقة بأخرى؛ انسب كل رقم لورقته بالاسم.
"""


# ==================== إعدادات النماذج وإعادة المحاولة ====================

# النموذج الافتراضي: أحدث نموذج **مستقر** (GA) وليس نسخة معاينة،
# لأن نماذج المعاينة (preview) هي مصدر الخطأ 503 UNAVAILABLE.
DEFAULT_MODEL = "gemini-3.8-flash"

# (المعرّف، الوصف المعروض، مستقر؟)
MODEL_CATALOG = [
    ("gemini-3.8-flash",       "الأحدث والمستقر — موصى به", True),
    ("gemini-3.6-flash",       "مستقر — بديل أخف", True),
    ("gemini-2.5-flash",       "مستقر — الجيل السابق", True),
    ("gemini-2.5-pro",         "مستقر — أبطأ وأدق", True),
    ("gemini-3-flash-preview", "معاينة — عرضة للازدحام (503)", False),
]

MODEL_OPTIONS = [m[0] for m in MODEL_CATALOG]
MODEL_LABELS = dict((m[0], m[1]) for m in MODEL_CATALOG)

# مستويات التفكير المدعومة لكل نموذج
# ملاحظة: gemini-3.8-flash لا يدعم MINIMAL — يُضبط تلقائياً إلى LOW
MODEL_THINKING_LEVELS = {
    "gemini-3.8-flash":       ("LOW", "MEDIUM", "HIGH"),
    "gemini-3.6-flash":       ("MINIMAL", "LOW", "MEDIUM", "HIGH"),
    "gemini-2.5-flash":       ("LOW", "MEDIUM", "HIGH"),
    "gemini-2.5-pro":         ("LOW", "MEDIUM", "HIGH"),
    "gemini-3-flash-preview": ("MINIMAL", "LOW", "MEDIUM", "HIGH"),
}

# سلسلة البدائل عند الازدحام — نماذج مستقرة فقط
FALLBACK_CHAIN = ["gemini-3.8-flash", "gemini-3.6-flash", "gemini-2.5-flash"]

LEVEL_ORDER = ["MINIMAL", "LOW", "MEDIUM", "HIGH"]

MAX_RETRIES = 4          # عدد المحاولات لكل نموذج
BASE_DELAY = 2.0         # ثانية، تتضاعف مع كل محاولة
TRANSIENT_CODES = (408, 429, 500, 502, 503, 504)


class GeminiUnavailable(Exception):
    """تعذّر الحصول على رد بعد استنفاد كل المحاولات والنماذج البديلة"""


def _error_code(exc) -> int:
    """استخراج كود HTTP من استثناء SDK أو من نص الرسالة"""
    for attr in ("code", "status_code", "http_status"):
        value = getattr(exc, attr, None)
        if isinstance(value, int):
            return value
    found = re.search(r"\b([45]\d{2})\b", str(exc))
    return int(found.group(1)) if found else 0


def _is_transient(exc) -> bool:
    """هل الخطأ مؤقت (ازدحام/حصة) ويستحق إعادة المحاولة؟"""
    if _error_code(exc) in TRANSIENT_CODES:
        return True
    text = str(exc).upper()
    keywords = ("UNAVAILABLE", "RESOURCE_EXHAUSTED", "OVERLOADED",
                "INTERNAL", "DEADLINE", "HIGH DEMAND", "TIMEOUT")
    return any(k in text for k in keywords)


def _clamp_thinking_level(model: str, level: str) -> str:
    """ضبط مستوى التفكير إلى أقرب مستوى يدعمه النموذج.

    يمنع خطأ 400 عند اختيار MINIMAL مع gemini-3.8-flash.
    """
    allowed = MODEL_THINKING_LEVELS.get(model, ("LOW", "MEDIUM", "HIGH"))
    if level in allowed:
        return level
    target = LEVEL_ORDER.index(level) if level in LEVEL_ORDER else len(LEVEL_ORDER) - 1
    return min(allowed, key=lambda a: abs(LEVEL_ORDER.index(a) - target))


def _build_config(model: str, thinking_level: str, system_instruction=None):
    """إعدادات التوليد — دفاعية ضد اختلاف إصدارات SDK"""
    level = _clamp_thinking_level(model, thinking_level)

    kwargs = {}
    if system_instruction:
        kwargs["system_instruction"] = system_instruction

    thinking = None
    try:
        thinking = types.ThinkingConfig(thinking_level=getattr(types.ThinkingLevel, level))
    except Exception:
        try:
            thinking = types.ThinkingConfig(thinking_level=level.lower())
        except Exception:
            thinking = None
    if thinking is not None:
        kwargs["thinking_config"] = thinking

    try:
        return types.GenerateContentConfig(**kwargs)
    except Exception:
        kwargs.pop("thinking_config", None)
        try:
            return types.GenerateContentConfig(**kwargs)
        except Exception:
            return None


def _model_chain() -> list:
    """النموذج المختار ثم البدائل المستقرة"""
    primary = st.session_state.get("model_name", DEFAULT_MODEL)
    chain = [primary]
    if st.session_state.get("allow_fallback", True):
        chain += [m for m in FALLBACK_CHAIN if m != primary]
    return chain


def generate_with_retry(client, content_parts, thinking_level: str, status=None,
                        system_instruction=None, stream_to=None) -> str:
    """استدعاء النموذج مع إعادة محاولة تصاعدية ثم التبديل إلى نموذج بديل.

    يعالج الخطأ 503 (ازدحام النموذج) بدل إسقاط التحليل بالكامل.
    يرفع GeminiUnavailable إذا فشلت كل المحاولات — ولا يُعيد نص خطأ
    حتى لا يتحوّل الخطأ إلى "تقرير" قابل للتحميل.

    إذا مُرّر stream_to (حاوية Streamlit) يُعرض النص تدريجياً أثناء توليده.
    """
    chain = _model_chain()
    primary = chain[0]
    last_error = None

    for model_name in chain:
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                config = _build_config(model_name, thinking_level, system_instruction)

                if stream_to is not None:
                    pieces = []
                    stream = client.models.generate_content_stream(
                        model=model_name, contents=content_parts, config=config)
                    for chunk in stream:
                        piece = getattr(chunk, "text", None)
                        if piece:
                            pieces.append(piece)
                            if len(pieces) % 8 == 0:
                                stream_to.markdown("".join(pieces))
                    text = "".join(pieces)
                    if text:
                        stream_to.markdown(text)
                    reason = "STREAM_EMPTY"
                else:
                    response = client.models.generate_content(
                        model=model_name, contents=content_parts, config=config)
                    text = getattr(response, "text", None)
                    reason = "غير معروف"
                    if not text:
                        try:
                            reason = str(response.candidates[0].finish_reason)
                        except Exception:
                            pass

                if not text:
                    raise GeminiUnavailable(
                        "النموذج %s أعاد رداً فارغاً (سبب الإنهاء: %s). "
                        "جرّب تقليل عدد المقالات أو مستوى التفكير." % (model_name, reason))

                if status is not None:
                    if model_name != primary:
                        status.info("ℹ️ تمّ التحليل بالنموذج البديل: %s" % model_name)
                    else:
                        status.empty()
                return text

            except GeminiUnavailable:
                raise
            except Exception as exc:
                last_error = exc
                if not _is_transient(exc):
                    raise
                if stream_to is not None:
                    stream_to.empty()
                if attempt == MAX_RETRIES:
                    if status is not None and len(chain) > 1:
                        status.warning(
                            "🔁 النموذج %s مزدحم — التبديل إلى نموذج بديل..." % model_name)
                    break
                delay = BASE_DELAY * (2 ** (attempt - 1)) + random.uniform(0, 1)
                if status is not None:
                    status.warning(
                        "⏳ النموذج %s مزدحم حالياً (503) — إعادة المحاولة %d/%d بعد %d ثانية..."
                        % (model_name, attempt, MAX_RETRIES - 1, round(delay)))
                time.sleep(delay)

    raise GeminiUnavailable(
        "تعذّر الوصول إلى أي نموذج بعد استنفاد كل المحاولات.\n\n"
        "آخر خطأ: %s" % last_error)


# تهيئة الحالة
if 'messages' not in st.session_state:
    st.session_state.messages = []
if 'analysis_result' not in st.session_state:
    st.session_state.analysis_result = None
if 'uploaded_files_content' not in st.session_state:
    st.session_state.uploaded_files_content = {}
if 'files_names' not in st.session_state:
    st.session_state.files_names = []
if 'client' not in st.session_state:
    st.session_state.client = None
if 'thinking_level' not in st.session_state:
    st.session_state.thinking_level = "HIGH"
if 'saved_api_key' not in st.session_state:
    st.session_state.saved_api_key = ""
if 'api_key_saved' not in st.session_state:
    st.session_state.api_key_saved = False
if 'model_name' not in st.session_state:
    st.session_state.model_name = DEFAULT_MODEL
if 'allow_fallback' not in st.session_state:
    st.session_state.allow_fallback = True


def _md_inline(text: str) -> str:
    """تنظيف رموز Markdown من نص سطري"""
    return text.replace("**", "").replace("`", "").strip()


def _set_rtl(paragraph):
    """ضبط اتجاه الفقرة والمقاطع من اليمين لليسار (مستوى الفقرة والمقطع)"""
    pPr = paragraph._p.get_or_add_pPr()
    bidi = OxmlElement("w:bidi")
    bidi.set(qn("w:val"), "1")
    pPr.append(bidi)
    for run in paragraph.runs:
        rPr = run._r.get_or_add_rPr()
        rtl = OxmlElement("w:rtl")
        rtl.set(qn("w:val"), "1")
        rPr.append(rtl)


def _add_table(doc, rows):
    """تحويل جدول Markdown إلى جدول Word حقيقي (بدل حذفه)"""
    cells = []
    for row in rows:
        line = row.strip().strip("|")
        parts = [_md_inline(p) for p in line.split("|")]
        # تجاهل سطر الفصل |---|---|
        if parts and all(set(p) <= set("-: ") and p for p in parts):
            continue
        cells.append(parts)

    cells = [c for c in cells if any(x for x in c)]
    if not cells:
        return

    width = max(len(c) for c in cells)
    table = doc.add_table(rows=0, cols=width)
    try:
        table.style = "Light Grid Accent 2"
    except Exception:
        table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.RIGHT

    for i, row_cells in enumerate(cells):
        cursor = table.add_row().cells
        for j in range(width):
            value = row_cells[j] if j < len(row_cells) else ""
            para = cursor[j].paragraphs[0]
            run = para.add_run(value)
            if i == 0:
                run.bold = True
            para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            _set_rtl(para)
    doc.add_paragraph()


def create_word_report(content: str, title: str = "تقرير التحليل") -> bytes:
    """إنشاء ملف Word من المحتوى مع دعم اللغة العربية والجداول"""
    doc = DocxDocument()

    # إعداد الصفحة للغة العربية (RTL)
    section = doc.sections[0]
    sectPr = section._sectPr
    bidi = OxmlElement("w:bidi")
    bidi.set(qn("w:val"), "1")
    sectPr.append(bidi)

    title_para = doc.add_heading("📊 منصة تحليل أوراق القياس الاقتصادي", level=0)
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    date_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    info_para = doc.add_paragraph()
    info_para.add_run("التاريخ: %s" % date_str).bold = True
    info_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _set_rtl(info_para)

    paper_para = doc.add_paragraph()
    paper_para.add_run("المقالة: %s" % title).bold = True
    paper_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    _set_rtl(paper_para)

    doc.add_paragraph()

    lines = content.split("\n")
    table_buffer = []
    in_code = False

    for line in lines:
        stripped = line.strip()

        # تجميع أسطر الجدول ثم إدراجه دفعة واحدة
        if stripped.startswith("|"):
            table_buffer.append(stripped)
            continue
        if table_buffer:
            _add_table(doc, table_buffer)
            table_buffer = []

        if stripped.startswith("```"):
            in_code = not in_code
            continue
        if not stripped:
            continue
        if in_code:
            para = doc.add_paragraph(line.rstrip())
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            continue
        if set(stripped) <= set("-─_=") and len(stripped) > 2:
            continue

        if stripped.startswith("### "):
            h = doc.add_heading(_md_inline(stripped[4:]), level=3)
            h.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            _set_rtl(h)
        elif stripped.startswith("## "):
            h = doc.add_heading(_md_inline(stripped[3:]), level=2)
            h.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            _set_rtl(h)
        elif stripped.startswith("# "):
            h = doc.add_heading(_md_inline(stripped[2:]), level=1)
            h.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            _set_rtl(h)
        elif stripped.startswith("- ") or stripped.startswith("* "):
            para = doc.add_paragraph(_md_inline(stripped[2:]), style="List Bullet")
            para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            _set_rtl(para)
        else:
            clean = _md_inline(stripped)
            if clean:
                para = doc.add_paragraph(clean)
                para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                _set_rtl(para)

    if table_buffer:
        _add_table(doc, table_buffer)

    doc.add_paragraph()
    doc.add_paragraph("─" * 50)
    footer = doc.add_paragraph()
    footer.add_run("تصميم وتطوير: د. مروان رودان | Dr. Merwan Roudane")
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def upload_file_to_gemini(client, file_content: bytes, filename: str):
    """رفع ملف إلى Gemini"""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(file_content)
        tmp_path = tmp.name

    try:
        uploaded_file = client.files.upload(file=tmp_path, config={"display_name": filename})
        return uploaded_file
    finally:
        os.unlink(tmp_path)


def _wait_until_active(client, file_ref, timeout: int = 180):
    """انتظار معالجة الملف على الخادم قبل استخدامه.

    ملفات PDF الكبيرة تبقى PROCESSING لثوانٍ؛ استخدامها قبل ACTIVE يُنتج خطأ.
    """
    deadline = time.time() + timeout
    while time.time() < deadline:
        state = getattr(file_ref, "state", None)
        state = getattr(state, "name", state)
        state = str(state) if state is not None else ""

        if state in ("ACTIVE", "", "None"):
            return file_ref
        if state == "FAILED":
            raise GeminiUnavailable(
                "فشل الخادم في معالجة الملف: %s"
                % getattr(file_ref, "display_name", "?"))

        time.sleep(2)
        try:
            file_ref = client.files.get(name=file_ref.name)
        except Exception:
            return file_ref
    return file_ref


def get_uploaded_files(client, files_content: dict, status=None) -> list:
    """رفع الملفات مرة واحدة فقط ثم إعادة استخدامها.

    النسخة السابقة كانت ترفع كل الـ PDF من جديد مع كل سؤال دردشة،
    أي 10 ملفات × 10 أسئلة = 100 عملية رفع لنفس البايتات.
    """
    cache = st.session_state.setdefault("uploaded_refs", {})
    refs = []

    for filename, data in files_content.items():
        key = "%s:%s" % (filename, hashlib.md5(data).hexdigest())
        ref = cache.get(key)
        if ref is None:
            if status is not None:
                status.info("📤 جارٍ رفع الملف: %s ..." % filename)
            try:
                ref = upload_file_to_gemini(client, data, filename)
                ref = _wait_until_active(client, ref)
                cache[key] = ref
            except GeminiUnavailable:
                raise
            except Exception as exc:
                if status is not None:
                    status.warning("⚠️ تعذر رفع الملف %s: %s" % (filename, exc))
                continue
        refs.append(ref)

    if status is not None:
        status.empty()
    return refs


def _instruction_for(files_content: dict) -> str:
    """تعليمات النظام: البرومبت الأساسي + قواعد الالتزام بالمصدر"""
    instruction = SYSTEM_PROMPT + GROUNDING_RULES
    if len(files_content) > 1:
        instruction += MULTI_PAPER_RULES
    return instruction


def analyze_with_gemini(client, files_content: dict, thinking_level: str,
                        question: str = None, status=None, stream_to=None) -> str:
    """تحليل المقالات"""
    uploaded_files = get_uploaded_files(client, files_content, status=status)
    if not uploaded_files:
        raise GeminiUnavailable("لم يتم رفع أي ملف بنجاح إلى الخادم")

    if question:
        task = (
            "بناءً على المقالات المرفقة، أجب على السؤال التالي بشكل مفصل ودقيق:\n\n"
            "%s\n\n"
            "قدم إجابة شاملة مع الإشارة للأقسام ذات الصلة في المقال." % question
        )
    else:
        task = (
            "قم بتحليل المقالات المرفقة تحليلاً شاملاً ومعمقاً وفقاً للهيكل "
            "والتعليمات المحددة في تعليمات النظام.\n\n"
            "ركز على:\n"
            "1. شرح النموذج/الاختبار الرئيسي بشكل مبسط وواضح\n"
            "2. استخراج جميع الفرضيات وشرحها\n"
            "3. تحديد متى يُستخدم ومتى لا يُستخدم بدقة\n"
            "4. استخراج معلومات المحاكاة وتنظيمها في جدول (بأرقام الورقة حرفياً)\n"
            "5. تقديم خطة تعلم عملية للباحث\n"
            "6. الإشارة للحالات الخاصة والتحذيرات\n\n"
            "الهدف: تمكين الباحث من فهم النموذج بشكل كامل وتطبيقه عملياً."
        )

    return generate_with_retry(
        client, uploaded_files + [task], thinking_level, status=status,
        system_instruction=_instruction_for(files_content), stream_to=stream_to)


def chat_with_gemini(client, files_content: dict, thinking_level: str,
                     messages: list, new_message: str, status=None) -> str:
    """الدردشة حول المقالات"""
    uploaded_files = get_uploaded_files(client, files_content, status=status)
    if not uploaded_files:
        raise GeminiUnavailable("لم يتم رفع أي ملف بنجاح إلى الخادم")

    # السؤال الجديد مضاف مسبقاً إلى messages — نستبعده من السياق لئلا يتكرر
    history = messages[-9:-1] if len(messages) > 1 else []
    chat_context = "\n\n".join([
        "%s: %s" % ("المستخدم" if m["role"] == "user" else "المساعد", m["content"])
        for m in history
    ])

    task = "بناءً على المقالات المرفقة"
    if chat_context:
        task += " والمحادثة السابقة:\n\n%s" % chat_context
    task += (
        "\n\nالسؤال الجديد: %s\n\n"
        "أجب بشكل دقيق ومفصل مع الإشارة للأجزاء المعنية في المقال. "
        "إذا كان السؤال عن نقطة لم تُفهم، اشرحها بطريقة مبسطة مع أمثلة." % new_message
    )

    return generate_with_retry(
        client, uploaded_files + [task], thinking_level, status=status,
        system_instruction=_instruction_for(files_content))


# ==================== الواجهة الرئيسية ====================

# الشريط الجانبي
with st.sidebar:
    st.markdown("""
    <div style='text-align: center; padding: 15px;'>
        <h3 style='color: white; margin: 0; font-size: 16px;'>📊 المنصة الشاملة</h3>
        <p style='color: #FFE4D6; font-size: 11px;'>القياس الاقتصادي</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    st.markdown("### 🔑 مفتاح API")
    
    # التحقق من وجود مفتاح محفوظ
    if st.session_state.api_key_saved and st.session_state.saved_api_key:
        st.success("✅ المفتاح محفوظ")
        st.markdown(f"<p style='color: #FFE4D6; font-size: 12px;'>المفتاح: ****{st.session_state.saved_api_key[-4:]}</p>", unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("🔄 تغيير", use_container_width=True, key="change_key"):
                st.session_state.api_key_saved = False
                st.session_state.saved_api_key = ""
                st.session_state.client = None
                st.rerun()
        
        api_key = st.session_state.saved_api_key
    else:
        api_key = st.text_input(
            "أدخل مفتاح Google API",
            type="password",
            help="احصل على مفتاح API من Google AI Studio",
            label_visibility="collapsed",
            placeholder="أدخل مفتاح API هنا..."
        )
        
        if api_key:
            col1, col2 = st.columns(2)
            with col1:
                if st.button("💾 حفظ", use_container_width=True, key="save_key"):
                    st.session_state.saved_api_key = api_key
                    st.session_state.api_key_saved = True
                    st.success("✅ تم الحفظ!")
                    st.rerun()
    
    st.markdown("### ⚙️ إعدادات النموذج")
    thinking_level = st.selectbox(
        "مستوى التفكير (Thinking Level)",
        ["HIGH", "MEDIUM", "LOW", "MINIMAL"],
        index=0,
        help="HIGH: تفكير عميق (افتراضي) | MEDIUM: متوسط | LOW: سريع | MINIMAL: أسرع"
    )

    model_name = st.selectbox(
        "النموذج (Model)",
        MODEL_OPTIONS,
        index=MODEL_OPTIONS.index(DEFAULT_MODEL),
        format_func=lambda m: "%s — %s" % (m, MODEL_LABELS.get(m, "")),
        help="النماذج المستقرة (GA) أقل عرضة للخطأ 503 من نماذج المعاينة (preview)"
    )
    st.session_state.model_name = model_name

    # بعض النماذج لا تدعم كل مستويات التفكير (مثلاً MINIMAL غير مدعوم في gemini-3.8-flash)
    _effective = _clamp_thinking_level(model_name, thinking_level)
    if _effective != thinking_level:
        st.caption("ℹ️ النموذج %s لا يدعم %s — سيُستخدم %s"
                   % (model_name, thinking_level, _effective))

    st.session_state.allow_fallback = st.checkbox(
        "التبديل التلقائي لنموذج بديل عند الازدحام",
        value=True,
        help="عند فشل النموذج الأساسي (503) تُعاد المحاولة ثم يُجرَّب نموذج آخر تلقائياً"
    )
    
    if api_key:
        if not st.session_state.api_key_saved:
            st.success("✅ تم إدخال المفتاح")
        try:
            # إنشاء العميل باستخدام SDK الجديد
            client = genai.Client(api_key=api_key)
            st.session_state.client = client
            st.session_state.thinking_level = thinking_level
        except Exception as e:
            st.error(f"خطأ في تكوين النموذج: {str(e)}")
    else:
        st.warning("⚠️ مطلوب مفتاح API")
    
    st.markdown("---")
    
    st.markdown("### 📁 رفع المقالات")
    uploaded_files = st.file_uploader(
        "اختر ملفات PDF",
        type=['pdf'],
        accept_multiple_files=True,
        help="الحد الأقصى: 10 مقالات",
        label_visibility="collapsed"
    )
    
    if uploaded_files:
        if len(uploaded_files) > 10:
            st.error("⚠️ الحد الأقصى 10 ملفات!")
            uploaded_files = uploaded_files[:10]
        
        st.success(f"✅ تم رفع {len(uploaded_files)} ملف")
        
        st.session_state.uploaded_files_content = {}
        st.session_state.files_names = []
        for file in uploaded_files:
            content = file.read()
            st.session_state.uploaded_files_content[file.name] = content
            st.session_state.files_names.append(file.name)
            file.seek(0)
        
        with st.expander("📋 الملفات المرفوعة", expanded=False):
            for name in st.session_state.files_names:
                st.write(f"📄 {name}")
    else:
        # إزالة الملفات من الرافع يجب أن تزيلها من التحليل أيضاً
        if st.session_state.uploaded_files_content:
            st.session_state.uploaded_files_content = {}
            st.session_state.files_names = []
            st.session_state.analysis_result = None
            st.session_state.messages = []
    
    st.markdown("---")
    
    st.markdown("""
    <div style='text-align: center; padding: 15px; background: rgba(255,255,255,0.1); border-radius: 10px;'>
        <p style='font-size: 11px; margin: 0; color: #FFE4D6;'>تصميم وتطوير</p>
        <p style='font-size: 14px; font-weight: bold; margin: 5px 0; color: white;'>د. مروان رودان</p>
        <p style='font-size: 10px; color: #FFE4D6;'>Dr. Marouane Roudan</p>
    </div>
    """, unsafe_allow_html=True)


# المحتوى الرئيسي
st.markdown("""
<div class='logo-container'>
    <p class='logo-title'>📊 المنصة الشاملة لشرح المقالات الأصلية والاختبارات في القياس الاقتصادي</p>
    <p class='logo-subtitle'>The Comprehensive Platform for Explaining Original Econometric Papers and Tests</p>
</div>
""", unsafe_allow_html=True)

# التبويبات
tab1, tab2, tab3 = st.tabs(["📑 تحليل المقالات", "💬 نافذة الدردشة", "📥 تحميل التقرير"])

# ==================== تبويب التحليل ====================
with tab1:
    st.markdown("### 🔍 التحليل الشامل للمقالات")
    
    if not api_key:
        st.warning("⚠️ يرجى إدخال مفتاح Google API في الشريط الجانبي للبدء")
    elif not st.session_state.uploaded_files_content:
        st.info("📁 يرجى رفع مقال واحد على الأقل من الشريط الجانبي")
    else:
        col1, col2 = st.columns([3, 1])
        
        with col1:
            st.markdown(f"""
            <div class='info-card'>
                <h4 style='color: #E07B39; margin: 0;'>📚 المقالات الجاهزة للتحليل</h4>
                <p style='font-size: 24px; font-weight: bold; color: #D4703A; margin: 10px 0;'>{len(st.session_state.uploaded_files_content)}</p>
                <p style='color: #888; font-size: 13px;'>اضغط على "بدء التحليل" للحصول على شرح شامل ومفصل</p>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            analyze_btn = st.button("🚀 بدء التحليل", use_container_width=True, type="primary")
        
        if analyze_btn:
            if st.session_state.client is None:
                st.error("❌ يرجى التأكد من إدخال مفتاح API صحيح")
            else:
                status = st.empty()
                stream_box = st.empty()
                with st.spinner("⏳ جاري تحليل المقالات... قد يستغرق هذا بضع دقائق"):
                    try:
                        result = analyze_with_gemini(
                            st.session_state.client,
                            st.session_state.uploaded_files_content,
                            st.session_state.thinking_level,
                            status=status,
                            stream_to=stream_box
                        )
                        stream_box.empty()
                        st.session_state.analysis_result = result
                    except GeminiUnavailable as e:
                        status.empty()
                        st.error("❌ تعذّر إتمام التحليل\n\n%s" % e)
                        st.info(
                            "💡 ماذا تفعل الآن؟\n\n"
                            "- أعد الضغط على «بدء التحليل» بعد دقيقة أو دقيقتين\n"
                            "- اختر نموذجاً مستقراً (gemini-2.5-flash) من الشريط الجانبي\n"
                            "- قلّل عدد المقالات أو اخفض مستوى التفكير إلى MEDIUM\n\n"
                            "الخطأ 503 يعني ازدحام خوادم النموذج، وليس خطأ في مفتاحك أو ملفاتك."
                        )
                    except Exception as e:
                        status.empty()
                        st.error("❌ حدث خطأ: %s" % e)
        
        if st.session_state.analysis_result:
            st.markdown("---")
            st.markdown("### 📊 نتائج التحليل")
            
            with st.container():
                st.markdown(f"""
                <div style='background: white; padding: 25px; border-radius: 15px; 
                            box-shadow: 0 4px 15px rgba(0,0,0,0.08); border: 1px solid #FFE4D6;'>
                """, unsafe_allow_html=True)
                
                st.markdown(st.session_state.analysis_result)
                
                st.markdown("</div>", unsafe_allow_html=True)

# ==================== تبويب الدردشة ====================
with tab2:
    st.markdown("### 💬 نافذة الدردشة التفاعلية")
    st.markdown("اطرح أسئلتك حول المقالات المرفوعة للحصول على توضيحات إضافية")
    
    if not api_key:
        st.warning("⚠️ يرجى إدخال مفتاح API أولاً")
    elif not st.session_state.uploaded_files_content:
        st.info("📁 يرجى رفع مقال واحد على الأقل")
    else:
        chat_container = st.container()
        
        with chat_container:
            for msg in st.session_state.messages:
                if msg['role'] == 'user':
                    st.markdown(
                        "<div class='chat-user'><strong>🧑‍🎓 أنت:</strong></div>",
                        unsafe_allow_html=True)
                else:
                    st.markdown(
                        "<div class='chat-assistant'><strong>🤖 المنصة:</strong></div>",
                        unsafe_allow_html=True)
                # يُعرض كـ Markdown حتى تظهر الجداول والمعادلات والعناوين
                st.markdown(msg['content'])
        
        st.markdown("---")
        
        col1, col2 = st.columns([5, 1])
        
        with col1:
            user_question = st.text_area(
                "اكتب سؤالك هنا",
                placeholder="مثال: ما هي الفرضية A3 بالتفصيل؟ أو اشرح لي نتائج المحاكاة...",
                height=80,
                label_visibility="collapsed"
            )
        
        with col2:
            send_btn = st.button("📤 إرسال", use_container_width=True, type="primary")
        
        st.markdown("**💡 أسئلة مقترحة:**")
        suggested_cols = st.columns(3)
        
        suggestions = [
            "ما هي الفرضيات الأساسية؟",
            "اشرح نتائج المحاكاة",
            "متى لا يُستخدم هذا النموذج؟"
        ]
        
        for i, suggestion in enumerate(suggestions):
            with suggested_cols[i]:
                if st.button(suggestion, key=f"sug_{i}", use_container_width=True):
                    user_question = suggestion
                    send_btn = True
        
        if send_btn and user_question:
            if st.session_state.client is None:
                st.error("❌ يرجى التأكد من إدخال مفتاح API صحيح")
            else:
                st.session_state.messages.append({
                    'role': 'user',
                    'content': user_question
                })
                
                status = st.empty()
                with st.spinner("⏳ جاري التفكير..."):
                    try:
                        response = chat_with_gemini(
                            st.session_state.client,
                            st.session_state.uploaded_files_content,
                            st.session_state.thinking_level,
                            st.session_state.messages,
                            user_question,
                            status=status
                        )

                        st.session_state.messages.append({
                            'role': 'assistant',
                            'content': response
                        })

                        st.rerun()
                    except GeminiUnavailable as e:
                        st.session_state.messages.pop()
                        status.empty()
                        st.error("❌ النموذج غير متاح حالياً\n\n%s" % e)
                    except Exception as e:
                        st.session_state.messages.pop()
                        status.empty()
                        st.error("❌ حدث خطأ: %s" % e)
        
        if st.session_state.messages:
            if st.button("🗑️ مسح المحادثة", type="secondary"):
                st.session_state.messages = []
                st.rerun()

# ==================== تبويب التحميل ====================
with tab3:
    st.markdown("### 📥 تحميل التقرير")
    
    if st.session_state.analysis_result:
        st.success("✅ التقرير جاهز للتحميل!")
        
        st.markdown(f"""
        <div class='info-card'>
            <h4 style='color: #E07B39;'>📄 معلومات التقرير</h4>
            <p><strong>عدد المقالات المحللة:</strong> {len(st.session_state.files_names)}</p>
            <p><strong>تاريخ التحليل:</strong> {datetime.now().strftime("%Y-%m-%d %H:%M")}</p>
            <p><strong>المقالات:</strong></p>
            <ul>
                {"".join([f"<li>{name}</li>" for name in st.session_state.files_names])}
            </ul>
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            try:
                word_bytes = create_word_report(
                    st.session_state.analysis_result,
                    ", ".join(st.session_state.files_names[:2]) + ("..." if len(st.session_state.files_names) > 2 else "")
                )
                
                st.download_button(
                    label="📥 تحميل Word",
                    data=word_bytes,
                    file_name=f"econometrics_report_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    type="primary"
                )
            except Exception as e:
                st.error(f"خطأ في إنشاء Word: {str(e)}")
        
        with col2:
            st.download_button(
                label="📄 تحميل TXT",
                data=st.session_state.analysis_result,
                file_name=f"econometrics_report_{datetime.now().strftime('%Y%m%d_%H%M')}.txt",
                mime="text/plain",
                use_container_width=True
            )
        
        with st.expander("👁️ معاينة التقرير", expanded=False):
            st.markdown(st.session_state.analysis_result)
    
    else:
        st.info("📊 قم بتحليل المقالات أولاً من تبويب 'تحليل المقالات' للحصول على تقرير قابل للتحميل")
        
        st.markdown("""
        <div class='info-card' style='text-align: center;'>
            <p style='font-size: 48px; margin: 20px 0;'>📑</p>
            <h4 style='color: #D4703A;'>لا يوجد تقرير بعد</h4>
            <p style='color: #888;'>ارفع مقالاتك وقم بتحليلها للحصول على تقرير شامل</p>
        </div>
        """, unsafe_allow_html=True)

# ==================== Footer ====================
st.markdown("---")
st.markdown("""
<div class='footer'>
    <p style='font-size: 15px; color: #E07B39; font-weight: bold;'>📊 المنصة الشاملة لشرح المقالات الأصلية والاختبارات في القياس الاقتصادي</p>
    <p class='designer-name'>تصميم وتطوير: د. مروان رودان | Dr. Marouane Roudan</p>
    <p style='font-size: 11px; color: #999; margin-top: 15px;'>© 2025</p>
</div>
""", unsafe_allow_html=True)
