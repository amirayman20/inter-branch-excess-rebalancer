# Inter-Branch Excess Rebalancer

<img src="https://raw.githubusercontent.com/amirayman20/inter-branch-excess-rebalancer/main/assets/logic_flow.png" width="100%">

---

## Badges

![Python](https://img.shields.io/badge/Python-3.8+-blue?style=for-the-badge&logo=python)
![Pandas](https://img.shields.io/badge/Pandas-Data_Analysis-green?style=for-the-badge&logo=pandas)
![Excel](https://img.shields.io/badge/Excel-Automated_Reports-217346?style=for-the-badge&logo=microsoft-excel)
![Supply Chain](https://img.shields.io/badge/Supply_Chain-Optimization-orange?style=for-the-badge)
![Business Logic](https://img.shields.io/badge/Algorithm-Inventory_Intelligence-purple?style=for-the-badge)

---

## Contact Me

![LinkedIn](https://img.shields.io/badge/LinkedIn-Amir_Ayman-0A66C2?style=for-the-badge&logo=linkedin)
![GitHub](https://img.shields.io/badge/GitHub-amirayman20-181717?style=for-the-badge&logo=github)
![Email](https://img.shields.io/badge/Email-amirayman20@gmail.com-red?style=for-the-badge&logo=gmail)

---

# المشكلة: مخزون محبوس بين الفروع

تخيل هذا السيناريو:

- صنف معين زائد في فرع A بكمية تكفي شهور طويلة
- نفس الصنف في فرع B شبه منتهي وعليه طلب عالي
- المستودع المركزي مالوش رصيد كافي يغطي الفرع B

النتيجة: فرع B بيطلب شراء جديد رغم إن الكمية فعليًا **موجودة فعلاً** في فرع A، لكن محبوسة بدون أي تحويل.

---

# الحل: Think Differently

المشروع دا مش بيراقب رصيد فرع لوحده... هو بيشوف الشبكة كلها (كل الفروع + المستودع) كنظام واحد، وبيحدد فين الكمية "الزيادة الحقيقية" وفين الاحتياج الحقيقي، وبيوصل بينهم تلقائيًا.

## الفكرة الأساسية
**نحدد EXCESS في كل فرع (بعد ترك تغطية كافية)، ونوصله لأعلى فرع NEED — لكن بس لو المستودع نفسه مش قادر يغطي الاحتياج.**

### Pipeline
Branch Sales + Stock → Excess/Need Calculation → Warehouse Check → Best-Match Transfer
---

# 1) Excess & Need Calculation

لكل فرع، بيتم حساب:

- **DAILY_CONSUMPTION** = المبيعات ÷ عدد أيام التقرير
- **EXCESS** = الرصيد الزائد عن تغطية MIN_COVERAGE_DAYS
- **NEED** = الكمية المطلوبة لتغطية NEED_DAYS القادمة

---

# 2) Warehouse Gate

التحويل بين الفروع بيحصل **فقط لو** رصيد المستودع المركزي للصنف ده ≤ حد أدنى معين (يعني المستودع مش قادر يغطي الاحتياج بنفسه).

لو المستودع عنده كمية كافية، التحويل بين الفروع مالوش لازمة.

---

# 3) Best-Match Selection + No Duplication

لكل صنف زايد في فرع معين:

- يتم فحص كل الفروع المحتملة اللي عندها احتياج فعلي لنفس الصنف
- يتم اختيار **فرع واحد فقط** — الأعلى احتياجًا (NEED)
- الصنف ده **مايتكررش** في أي تقرير تحويل تاني، حتى لو كان زايد في فرع تاني كمان

---

# مثال عملي

- صنف X زايد بمقدار 40 وحدة في فرع A
- نفس الصنف محتاج 25 وحدة في فرع B، و15 وحدة في فرع C
- المستودع عنده رصيد صفر من الصنف

**القرار:** يتحول 25 وحدة من A إلى B (الأعلى احتياجًا)، ويُستبعد الصنف من أي تحويل تالٍ.

---

# Excel Output

كل تقرير تحويل (Source → Target) بيتولد كملف Excel منفصل، وفيه:

| Column | Description |
|--------|--------------|
| ITEM_CODE | كود الصنف |
| ITEM_NAME | اسم الصنف |
| TRANSFER_QTY | الكمية المقترح تحويلها |
| EXPIRE_DATE | تاريخ الصلاحية |
| REASON | سبب التحويل (مصفّر / طلب عالي / زيادة في المصدر...) |

---

# Tech Stack

- Python 3.8+
- Pandas
- Excel Automation (openpyxl)

---

# Usage

## 1) Prepare Data

- **warehouse.xlsx** → ITEM_CODE, ST_BAL
- **branch files (×N)** → ITEM_CODE, ITEM_NAME, EXPIRE_DATE, SALES_QTY, STORE_NAME, BALANCE

لو مش عندك بيانات حقيقية تجرب بيها، استخدم `generate_synthetic_data.py` لتوليد بيانات وهمية بنفس الهيكل تلقائيًا.

## 2) Update Paths

```python
warehouse_path = r"C:\Users\YourName\Desktop\warehouse.xlsx"
branch_files = [
    r"C:\Users\YourName\Desktop\branch_1.xlsx",
    r"C:\Users\YourName\Desktop\branch_2.xlsx",
    ...
]
```

## 3) Run

```bash
python main_transfer_script.py
```

## 4) Output

مجلد `Branch_Rebalancing_Results/` فيه ملف Excel لكل زوج (مصدر → مستهدف) فيه التحويلات المقترحة فقط.

---

# Settings

```python
MIN_WAREHOUSE_QTY  = 12    # لو رصيد المستودع للصنف ≤ 12، يبقى المستودع مش قادر يغطي → ندخل في التحويل بين الفروع
MIN_COVERAGE_DAYS  = 60    # نسيب في الفرع المصدر تغطية تكفيه 60 يوم، والباقي يُعتبر EXCESS قابل للتحويل
NEED_DAYS          = 30    # نحسب احتياج الفرع المستهدف على أساس تغطية 30 يوم قدام
EXPIRE_LIMIT_DAYS  = 120   # نستبعد أي صنف باقي على انتهاء صلاحيته أقل من 120 يوم (مايصلحش للتحويل)
QUARTER_DAYS       = 120   # عدد أيام فترة المبيعات المستخدمة في حساب الاستهلاك اليومي (تقريبًا ربع سنة)
```

---

# FAQ

### ليه التحويل مش بيحصل لكل الفروع المحتاجة؟
لأن نفس الصنف بيتحول لفرع واحد فقط (الأعلى احتياجًا)، لتجنب تشتيت الكمية الزائدة وتكرارها في أكتر من تقرير.

### ينفع لشركات غير صيدليات؟
نعم — أي شركة عندها شبكة فروع + مستودع مركزي ومبيعات متفاوتة بين الفروع.

---

**Last Updated:** June 2026
