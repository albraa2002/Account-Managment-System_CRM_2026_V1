# 🏢 Al Arrab Agency — Account Management System (AMMS)

نظام إدارة قسم Account Management لشركة العراب للتسويق الرقمي.

---

## 📋 نظرة عامة

سيستم داخلي متكامل لإدارة العملاء، المتابعات، التقارير، التجديدات، والعقود — مبني على Google Sheets كـ Database و Apps Script كـ Backend.

---

## ⚙️ التقنيات المستخدمة

| Layer | Technology |
|---|---|
| Frontend | HTML + CSS + JavaScript (Vanilla) |
| Backend | Google Apps Script |
| Database | Google Sheets |
| Charts | Chart.js |
| Hosting | GitHub Pages |

---

## 🗂️ محتوى الملفات

```
AMMS/
├── index.html      → Dashboard (Frontend كامل)
├── Code.gs         → Backend API (CRUD + Automation)
├── Setup.gs        → Sheet Builder (شغّله مرة واحدة)
└── README.md       → هذا الملف
```

---

## 🚀 طريقة التشغيل

### 1. إعداد Google Sheets

1. افتح Google Sheets جديد
2. من القائمة: **Extensions → Apps Script**
3. أنشئ ملفين: `Code.gs` و `Setup.gs`
4. انسخ الكود من الملفات المرفقة
5. شغّل `setupAll()` مرة واحدة — هيبني كل الـ Sheets تلقائياً

### 2. Deploy كـ Web App

1. اضغط **Deploy → New Deployment**
2. النوع: **Web App**
3. Execute as: **Me**
4. Who has access: **Anyone**
5. اضغط **Deploy** وانسخ الـ URL

### 3. ربط الـ Dashboard

1. افتح `index.html` في المتصفح
2. الصق الـ Web App URL في خانة الـ Connect
3. اضغط **Connect** ✅

---

## 📊 الصفحات والمميزات

| الصفحة | الوصف |
|---|---|
| **Dashboard** | نظرة عامة — KPIs + Charts + Alerts |
| **Clients** | إدارة العملاء — Add / Edit / Filter / Search |
| **Reports** | متابعة التقارير الأسبوعية والشهرية |
| **Follow-ups** | جدولة المتابعات — Phone / WhatsApp / Meeting |
| **Renewals** | تتبع تجديدات العقود مع تنبيهات |
| **Contracts** | إدارة العقود والمدفوعات والإيرادات |
| **Performance** | أداء الفريق — Scorecard per Manager |
| **Team** | إدارة أعضاء الفريق |

---

## 👥 هيكل الفريق

| الدور | عدد العملاء |
|---|---|
| Head of Account | 0 (إشراف) |
| Team Leader | 6 |
| Senior Account Manager | 10 |
| Account Manager | 10 |

---

## 🔧 الـ Sheets المُنشأة تلقائياً

- **Team** — بيانات أعضاء الفريق
- **Clients** — بيانات العملاء الكاملة
- **Contracts** — العقود والمدفوعات
- **Reports** — التقارير الأسبوعية والشهرية
- **FollowUps** — سجل المتابعات
- **ActivityLog** — سجل كل العمليات

---

## ⚡ الـ Automation

- **Daily Trigger (8 AM):** تحديث تلقائي لحالات التقارير والمتابعات والمدفوعات
- **Email Summary:** ملخص يومي يُرسل تلقائياً لـ Head of Account

---

## 🛠️ التطوير المستقبلي

- [ ] Role-based access (لكل مستخدم صلاحياته)
- [ ] Looker Studio integration
- [ ] Mobile app (PWA)
- [ ] WhatsApp reminders automation

---

## 📞 التواصل

**Al Arrab Digital Marketing Agency**  
Built by the Data & Systems Team
