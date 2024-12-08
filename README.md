<!-- markdownlint-disable MD030 -->

<img width="100%" src="https://github.com/user-attachments/assets/a0e2b28d-c0c2-4e41-9da8-91a8b5102e24" alt="Maximo Project Banner"> </p>

# Maximo Project

[![Python Version](https://img.shields.io/badge/python-3.10%2B-brightgreen)](https://www.python.org/)
[![Node.js](https://img.shields.io/badge/node.js-22.12.0-yellowgreen)](https://nodejs.org/)
[![Stars](https://img.shields.io/github/stars/naravid19/maximo_project?style=social)](https://github.com/naravid19/maximo_project)

โปรเจกต์ที่รวมการทำงานระหว่าง **Node.js** และ **Python** โดยมี **Backend** ทำงานด้วย **Django 5.1** และ **Frontend** ใช้ **Tailwind CSS** และ **Flowbite** พร้อมฟีเจอร์การตรวจสอบข้อมูล การสร้าง Template สำหรับ **MxLoader** และการจัดการ **JB-PM Plan**

## 🚀 **ฟีเจอร์หลัก**

- **📝 ตรวจสอบข้อมูล (Data Validation)**  
  ตรวจสอบความถูกต้องของข้อมูล ช่วยลดข้อผิดพลาดในการอัปโหลดเข้าสู่ระบบ

- **📄 สร้าง Template สำหรับ MxLoader**  
  รองรับการสร้างและจัดการ Template สำหรับ **MxLoader** เพื่อการนำเข้าข้อมูลเข้าสู่ระบบ Maximo

- **🗂️ จัดทำแผน JB-PM (Job Plan Management)**  
  สร้างและจัดการแผนงาน **JB-PM Plan** ได้อย่างมีประสิทธิภาพ

- **⚙️ Backend ทำงานด้วย Django 5.1**  
  พัฒนาเว็บแอปพลิเคชันที่ปลอดภัยและมีประสิทธิภาพสูง

- **🎨 Frontend ใช้ Tailwind CSS และ Flowbite**  
  ออกแบบ UI อย่างรวดเร็วด้วย Tailwind CSS และใช้ Component สำเร็จรูปจาก Flowbite เพื่อเพิ่มความสะดวกในการพัฒนา

## 🛠️ **สิ่งที่ต้องเตรียมก่อนติดตั้ง**

โปรดติดตั้งซอฟต์แวร์ต่อไปนี้ก่อนเริ่มใช้งาน:

1. **Node.js** (เวอร์ชัน 22.12.0) - [ดาวน์โหลด Node.js](https://nodejs.org/)  
2. **Python** (เวอร์ชัน 3.10 ขึ้นไป) - [ดาวน์โหลด Python](https://www.python.org/)

## 📦 **ขั้นตอนการติดตั้ง**

### 1. โคลนโปรเจกต์จาก GitHub

```bash
git clone https://github.com/naravid19/maximo_project.git
cd maximo_project
```

### 2. ติดตั้งไลบรารีของ Python

```bash
pip install -r requirements.txt
```

### 3. ติดตั้ง Tailwind CSS

```bash
npm install -D tailwindcss
```

### 4. ติดตั้ง Flowbite

```bash
npm install flowbite
```

## 💻 **การคอมไพล์ CSS ด้วย Tailwind CSS**

### คอมไพล์ CSS หนึ่งครั้ง

```bash
npm run build
```

### คอมไพล์ CSS แบบเรียลไทม์

```bash
npm run watch
```

## 💻 **การรัน Backend (Django)**

```bash
python manage.py runserver
```

เปิดใช้งาน Backend ได้ที่ [http://localhost:8000](http://localhost:8000)

## 📂 **โครงสร้างโปรเจกต์**

```
maximo_project/
│-- backend/
│   └── manage.py
│   └── requirements.txt
│   └── maximo_app/
│       └── templates/
│           └── maximo_app/
│               └── upload_form.html
│-- frontend/
│   └── package.json
│   └── tailwind.config.js
│-- static/
│   └── css/
│       └── styles.css
└── README.md
```

## 📖 **เอกสารประกอบ**

- **Django Documentation**: [https://docs.djangoproject.com/](https://docs.djangoproject.com/)  
- **Tailwind CSS Documentation**: [https://tailwindcss.com/docs](https://tailwindcss.com/docs)  
- **Flowbite Documentation**: [https://flowbite.com/docs](https://flowbite.com/docs)

## 🤝 **การมีส่วนร่วม (Contributing)**

ยินดีต้อนรับทุกการมีส่วนร่วม! โปรด Fork โปรเจกต์และส่ง Pull Request พร้อมคำอธิบายที่ชัดเจน

## 🌟 **ให้กำลังใจโดยการกดดาว ⭐ บน GitHub**

[![GitHub star chart](https://img.shields.io/github/stars/naravid19/maximo_project?style=social)](https://github.com/naravid19/maximo_project)
