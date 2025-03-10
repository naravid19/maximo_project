<!-- markdownlint-disable MD030 -->
<p align="center">
  <img src="https://github.com/user-attachments/assets/54550f9d-5f29-4d87-a573-c95b53a637b0" alt="Maximo Project Banner" style="width:100%">
</p>

# Maximo Project

**Maximo Project** คือโปรแกรมตรวจสอบและจัดทำข้อมูลงานบำรุงรักษาแบบหยุดตามวาระ (Shutdown Maintenance) ในรูปแบบใหม่  
โดยผสานระบบบริหารจัดการงานบำรุงรักษาด้วยคอมพิวเตอร์ (CMMS) เพื่อให้กระบวนการจัดทำข้อมูลมีความถูกต้อง โปร่งใส และรวดเร็ว

---

## ⚙️ เกริ่นนำ

โปรเจกต์นี้ออกแบบมาเพื่อแก้ปัญหาความซับซ้อนและข้อผิดพลาดในกระบวนการจัดทำ:
- **Job Plan Task** (ใบงานงาน)
- **Job Plan Labor** (ใบงานค่าแรง)
- **PM Plan** (แผนบำรุงรักษาตามวาระ)

โดยข้อมูลจะถูกจัดทำเป็นไฟล์ Excel Macro-Enabled Workbook ที่สามารถใช้งานร่วมกับ **MxLoader (v8.4.2)** ได้อย่างมีประสิทธิภาพ  
ซึ่งช่วยสนับสนุนการดำเนินงานในโรงไฟฟ้าให้มีความปลอดภัยและพร้อมจ่ายไฟฟ้าได้รวดเร็ว

---

## 🌟 ฟีเจอร์หลัก

- **Data Validation**  
  ตรวจสอบและกลั่นกรองข้อมูลก่อนสร้างใบงาน ลดข้อผิดพลาดและเพิ่มความถูกต้อง

- **MxLoader Template Generation**  
  สร้างไฟล์ Excel Macro-Enabled สำหรับนำเข้าข้อมูลสู่ระบบ Maximo อย่างครบถ้วน

- **Job Plan & PM Plan Management**  
  จัดกลุ่มใบงาน (Job Plan Task, Job Plan Labor) และสร้าง PM Plan ด้วยวิธีที่มีประสิทธิภาพ

- **Responsive UI Development**  
  พัฒนา Frontend ทันสมัยด้วย **Tailwind CSS**, **Flowbite** และ **jQuery**

- **Robust Backend**  
  ใช้ **Django 5.1** ผสานกับ **Pandas, NumPy, OpenPyXL** เพื่อประมวลผลข้อมูล Excel และจัดการไฟล์

- **Background Task Processing**  
  รันงาน background (เช่น การลบไฟล์เก่า) 

---

## 📋 สิ่งที่ต้องเตรียม

- **Node.js** (เวอร์ชัน LTS)  
  [ดาวน์โหลด Node.js](https://nodejs.org/)

- **Python** (เวอร์ชัน 3.10 ขึ้นไป)  
  [ดาวน์โหลด Python](https://www.python.org/)

---

## 🔧 การติดตั้ง

### 1. โคลนโปรเจกต์

```bash
git clone https://github.com/naravid19/maximo_project.git
cd maximo_project
```

### 2. ติดตั้งไลบรารี Python

```bash
pip install -r requirements.txt
```

### 3. ติดตั้ง Dependencies Node.js

```bash
npm install
```

### 4. คอมไพล์ CSS ด้วย Tailwind CSS

- **คอมไพล์ครั้งเดียว**

  ```bash
  npm run build
  ```

- **คอมไพล์แบบเรียลไทม์**

  ```bash
  npm run watch
  ```

> **หมายเหตุ:** สำหรับ production ควรรันคำสั่ง  
> ```bash
> python manage.py collectstatic --noinput
> ```  
> เพื่อรวบรวมไฟล์ static ไว้ในโฟลเดอร์ที่กำหนด

---

## 🚀 การใช้งาน

### รัน Backend (Django)

```bash
python manage.py runserver
```

เปิดเว็บแอปที่ [http://localhost:8000](http://localhost:8000)

### Background Task Processing

เพื่อให้ task ทำงาน (เช่น ลบไฟล์เก่า) ต้องรันคำสั่ง:

```bash
python manage.py process_tasks
```

ใน production คุณสามารถใช้ process manager (เช่น Supervisor หรือ systemd) เพื่อให้คำสั่งนี้ทำงานตลอดเวลา

## 🗂️ โครงสร้างโปรเจกต์

```
maximo_project/
│
├── manage.py
├── requirements.txt
│
├── maximo_app/
│   ├── models.py              # โครงสร้างฐานข้อมูล
│   ├── views.py               # ฟังก์ชันจัดการและประมวลผลข้อมูล
│   ├── forms.py               # ฟอร์มรับข้อมูลผู้ใช้
│   ├── urls.py                # URL และ API สำหรับ AJAX
│   ├── tasks.py               # งาน background (เช่น ลบไฟล์เก่า)
│   └── templates/
│       └── maximo_app/
│           └── upload_form.html
│
├── static/
│   ├── js/                    # ไฟล์ JavaScript (all-code.js, ajax-*.js, grouping.js, upload-loading.js, etc.)
│   └── src/                   # Tailwind CSS input/output
│
├── package.json               # การตั้งค่า Node.js และ Tailwind CSS
├── tailwind.config.js         # กำหนดค่า Tailwind CSS
└── README.md
```

---

## 🤝 การมีส่วนร่วม

นักพัฒนาที่สนใจสามารถ:
- **Fork** โปรเจกต์นี้และส่ง Pull Request เพื่อเพิ่มฟีเจอร์หรือแก้ไขบั๊ก
- เปิด **Issue** หากพบปัญหาหรือมีข้อเสนอแนะ

---

## 📜 License

โปรเจกต์นี้ใช้ **MIT License**  
ดูรายละเอียดเพิ่มเติมในไฟล์ [LICENSE](LICENSE)

---

## 📚 เอกสารอ้างอิง

- **Django Documentation:** [https://docs.djangoproject.com/](https://docs.djangoproject.com/)
- **Tailwind CSS Documentation:** [https://tailwindcss.com/docs](https://tailwindcss.com/docs)
- **Flowbite Documentation:** [https://flowbite.com/docs](https://flowbite.com/docs)
- **Pandas Documentation:** [https://pandas.pydata.org/docs/](https://pandas.pydata.org/docs/)

---

## 💡 สรุป

Maximo Project เป็นโซลูชันครบวงจรสำหรับการตรวจสอบและจัดทำข้อมูลงานบำรุงรักษา  
ที่ช่วยลดข้อผิดพลาด เพิ่มความโปร่งใสและความรวดเร็วในกระบวนการดำเนินงาน  
เพื่อให้โรงไฟฟ้าสามารถดำเนินงานได้อย่างราบรื่นและปลอดภัย

---

*สร้างสรรค์โดยทีม Maximo Project – เปลี่ยนวิธีจัดการงานบำรุงรักษาเพื่ออนาคตที่ดีกว่า*

---

หากมีคำแนะนำเพิ่มเติมหรือปรับปรุงในส่วนใด แจ้งได้เลยครับ!
