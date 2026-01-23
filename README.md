<!-- Improved compatibility of back to top link -->

<a id="readme-top"></a>

<!-- PROJECT SHIELDS -->

[![Contributors][contributors-shield]][contributors-url]
[![Forks][forks-shield]][forks-url]
[![Stargazers][stars-shield]][stars-url]
[![Issues][issues-shield]][issues-url]
[![MIT License][license-shield]][license-url]
[![LinkedIn][linkedin-shield]][linkedin-url]

<!-- PROJECT LOGO -->
<br />
<div align="center">
  <a href="https://github.com/naravid19/maximo_project">
    <img src="https://github.com/user-attachments/assets/54550f9d-5f29-4d87-a573-c95b53a637b0" alt="Maximo Project Logo" width="100%">
  </a>

  <h3 align="center">Maximo Project</h3>

  <p align="center">
    ระบบบริหารจัดการและจัดทำข้อมูลงานบำรุงรักษาแบบหยุดตามวาระสำหรับโรงไฟฟ้า (Planned Outage Management System)
    <br />
    <a href="#getting-started"><strong>เริ่มต้นใช้งาน »</strong></a>
    <br />
    <br />
    <a href="#usage">ดูตัวอย่าง</a>
    &middot;
    <a href="https://github.com/naravid19/maximo_project/issues">แจ้งปัญหา</a>
    &middot;
    <a href="https://github.com/naravid19/maximo_project/issues">ขอฟีเจอร์ใหม่</a>
  </p>
</div>

<!-- TABLE OF CONTENTS -->
<details>
  <summary>สารบัญ</summary>
  <ol>
    <li>
      <a href="#about-the-project">เกี่ยวกับโปรเจกต์</a>
      <ul>
        <li><a href="#built-with">เทคโนโลยีที่ใช้</a></li>
      </ul>
    </li>
    <li>
      <a href="#getting-started">การเริ่มต้นใช้งาน</a>
      <ul>
        <li><a href="#prerequisites">สิ่งที่ต้องเตรียม</a></li>
        <li><a href="#installation">การติดตั้ง</a></li>
      </ul>
    </li>
    <li><a href="#usage">การใช้งาน</a></li>
    <li><a href="#roadmap">แผนในอนาคต</a></li>
    <li><a href="#changelog">ประวัติการเปลี่ยนแปลง</a></li>
    <li><a href="#contributing">การมีส่วนร่วม</a></li>
    <li><a href="#license">สัญญาอนุญาต</a></li>
    <li><a href="#contact">ติดต่อ</a></li>
    <li><a href="#acknowledgments">กิตติกรรมประกาศ</a></li>
  </ol>
</details>

<!-- ABOUT THE PROJECT -->

## ℹ️ เกี่ยวกับโปรเจกต์

**Maximo Project** คือโซลูชันนวัตกรรมที่ออกแบบมาเพื่อปฏิวัติกระบวนการจัดทำข้อมูลงานบำรุงรักษาแบบหยุดตามวาระ (Planned Outage) เชื่อมต่อช่องว่างระหว่างการวางแผนงานและการนำเข้าข้อมูลสู่ระบบ CMMS (Maximo) ช่วยให้วิศวกรและผู้เกี่ยวข้องทำงานได้รวดเร็ว แม่นยำ และมีประสิทธิภาพสูงสุด

ระบบนี้ถูกพัฒนาเพื่อแก้ปัญหาความซับซ้อนในการจัดการ:

- 📄 **Job Plan Task** (ใบงานแผนงาน)
- 👷 **Job Plan Labor** (ใบงานค่าแรงและกำลังคน)
- 📅 **PM Plan** (แผนบำรุงรักษาเชิงป้องกันตามวาระ)

โดยผลลัพธ์ที่ได้คือไฟล์ Excel Macro-Enabled ที่พร้อมใช้งานกับ **MxLoader (v8.4.2)** ทันที ลดเวลาการเตรียมข้อมูลจากเดิมที่ใช้เวลาหลายวันเหลือเพียงไม่กี่นาที

<p align="right">(<a href="#readme-top">กลับไปด้านบน</a>)</p>

### 🛠️ เทคโนโลยีที่ใช้

โปรเจกต์นี้พัฒนาด้วยเทคโนโลยีที่ทันสมัยและมีเสถียรภาพ เพื่อให้มั่นใจในประสิทธิภาพและการรองรับการขยายตัวในอนาคต

- [![Django][Django]][Django-url] - Web Framework ประสิทธิภาพสูง
- [![TailwindCSS][TailwindCSS]][TailwindCSS-url] - Utility-first CSS framework สำหรับ UI ที่สวยงาม
- [![Flowbite][Flowbite]][Flowbite-url] - UI Component library
- [![Pandas][Pandas]][Pandas-url] - Library ประมวลผลและวิเคราะห์ข้อมูล Excel
- [![JQuery][JQuery.com]][JQuery-url] - จัดการ DOM และ AJAX interactions

<p align="right">(<a href="#readme-top">กลับไปด้านบน</a>)</p>

<!-- GETTING STARTED -->

## 🚀 การเริ่มต้นใช้งาน

ทำตามขั้นตอนด้านล่างเพื่อติดตั้งและรันโปรเจกต์บนเครื่องของคุณสำหรับการพัฒนาหรือทดสอบ

### 📋 สิ่งที่ต้องเตรียม

เพื่อให้แน่ใจว่าโปรแกรมจะทำงานได้อย่างสมบูรณ์ โปรดตรวจสอบว่าเครื่องของคุณมีซอฟต์แวร์ต่อไปนี้:

- **Python (3.12+)**: [ติดตั้ง Python](https://www.python.org/downloads/)
- **Node.js (LTS version)**: [ติดตั้ง Node.js](https://nodejs.org/)
- **Git**: [ติดตั้ง Git](https://git-scm.com/)

### 🔧 การติดตั้ง

1.  **Clone โปรเจกต์**

    ```sh
    git clone https://github.com/naravid19/maximo_project.git
    cd maximo_project
    ```

2.  **ตั้งค่า Python Environment**

    ```sh
    # สร้าง Virtual Environment
    python -m venv venv

    # Active Environment
    # Windows:
    .\venv\Scripts\activate
    # macOS/Linux:
    source venv/bin/activate
    ```

3.  **ติดตั้ง Python Dependencies**

    ```sh
    pip install -r requirements.txt
    ```

4.  **ติดตั้ง Node.js Dependencies และ Build CSS**

    ```sh
    npm install
    npm run build
    ```

5.  **ตั้งค่า Environment Variables**
    เปลี่ยนชื่อไฟล์ `.env.example` เป็น `.env` และแก้ไขค่า `DJANGO_SECRET_KEY` ให้ปลอดภัย

    ```sh
    cp .env.example .env
    ```

6.  **เริ่มต้นใช้งาน**
    ```sh
    python manage.py migrate
    python manage.py runserver
    ```
    เข้าใช้งานได้ที่ [http://localhost:8000](http://localhost:8000)

<p align="right">(<a href="#readme-top">กลับไปด้านบน</a>)</p>

<!-- USAGE EXAMPLES -->

## 💡 การใช้งาน

ระบบออกแบบมาให้ใช้งานง่ายใน 3 ขั้นตอนหลัก:

1.  **Upload & Config**: อัปโหลดไฟล์แผนงาน (Schedule) และไฟล์สถานที่ (Location) พร้อมระบุปีและเงื่อนไขการกรอง
2.  **Validate**: ระบบจะตรวจสอบความถูกต้องของข้อมูล (Data Validation) และแจ้งเตือนหากพบข้อผิดพลาด
3.  **Download**: ดาวน์โหลดไฟล์ Excel ที่ผ่านการประมวลผลแล้ว เพื่อนำไปใช้กับ MxLoader ต่อไป

### Background Tasks

ระบบมีระบบจัดการงานเบื้องหลังเพื่อทำความสะอาดไฟล์ชั่วคราว สามารถรันได้ด้วยคำสั่ง:

```sh
python manage.py process_tasks
```

<p align="right">(<a href="#readme-top">กลับไปด้านบน</a>)</p>

<!-- ROADMAP -->

## 🗺️ แผนในอนาคต

- [x] ระบบสร้าง Template สำหรับ MxLoader
- [x] ระบบตรวจสอบความถูกต้องของข้อมูล (Validation) และแจ้งเตือนข้อผิดพลาด
- [x] UI/UX ปรับปรุงใหม่ พร้อม Responsive Design
- [ ] เพิ่ม Dashboard สรุปภาพรวมสถิติข้อมูล
- [ ] รองรับการ Export ข้อมูลในรูปแบบ JSON API
- [ ] ระบบจัดการสิทธิ์ผู้ใช้งาน (Authentication & Authorization)
- [ ] รองรับหลายภาษา (Multi-language Support)

<p align="right">(<a href="#readme-top">กลับไปด้านบน</a>)</p>

<!-- CHANGELOG -->

## 📅 ประวัติการเปลี่ยนแปลง

ดูรายละเอียดการเปลี่ยนแปลงทั้งหมดได้ที่ไฟล์ [CHANGELOG.md](CHANGELOG.md)

<p align="right">(<a href="#readme-top">กลับไปด้านบน</a>)</p>

<!-- CONTRIBUTING -->

## 🤝 การมีส่วนร่วม

การมีส่วนร่วมคือกุญแจสำคัญที่ทำให้ชุมชนโอเพนซอร์สเติบโต เรายินดีต้อนรับทุกคำแนะนำและการช่วยเหลือ:

1.  **Fork** โปรเจกต์
2.  สร้าง Branch สำหรับฟีเจอร์ของคุณ (`git checkout -b feature/AmazingFeature`)
3.  Commit การเปลี่ยนแปลง (`git commit -m 'Add some AmazingFeature'`)
4.  Push ไปยัง Branch (`git push origin feature/AmazingFeature`)
5.  เปิด **Pull Request**

<p align="right">(<a href="#readme-top">กลับไปด้านบน</a>)</p>

<!-- LICENSE -->

## 📜 สัญญาอนุญาต

โปรเจกต์นี้เผยแพร่ภายใต้สัญญาอนุญาต **MIT License** ดูรายละเอียดเพิ่มเติมในไฟล์ `LICENSE`

<p align="right">(<a href="#readme-top">กลับไปด้านบน</a>)</p>

<!-- CONTACT -->

## 📧 ติดต่อ

Naravid - [github.com/naravid19](https://github.com/naravid19)

Project Link: [https://github.com/naravid19/maximo_project](https://github.com/naravid19/maximo_project)

<p align="right">(<a href="#readme-top">กลับไปด้านบน</a>)</p>

<!-- ACKNOWLEDGMENTS -->

## 🙏 กิตติกรรมประกาศ

ขอขอบคุณเครื่องมือและไลบรารีดีๆ ที่ช่วยให้โปรเจกต์นี้เกิดขึ้นได้:

- [Django](https://www.djangoproject.com/)
- [Tailwind CSS](https://tailwindcss.com)
- [Flowbite](https://flowbite.com)
- [MxLoader](https://www.mro.com/)
- [Img Shields](https://shields.io)

<p align="right">(<a href="#readme-top">กลับไปด้านบน</a>)</p>

<!-- MARKDOWN LINKS & IMAGES -->

[contributors-shield]: https://img.shields.io/github/contributors/naravid19/maximo_project.svg?style=for-the-badge
[contributors-url]: https://github.com/naravid19/maximo_project/graphs/contributors
[forks-shield]: https://img.shields.io/github/forks/naravid19/maximo_project.svg?style=for-the-badge
[forks-url]: https://github.com/naravid19/maximo_project/network/members
[stars-shield]: https://img.shields.io/github/stars/naravid19/maximo_project.svg?style=for-the-badge
[stars-url]: https://github.com/naravid19/maximo_project/stargazers
[issues-shield]: https://img.shields.io/github/issues/naravid19/maximo_project.svg?style=for-the-badge
[issues-url]: https://github.com/naravid19/maximo_project/issues
[license-shield]: https://img.shields.io/github/license/naravid19/maximo_project.svg?style=for-the-badge
[license-url]: https://github.com/naravid19/maximo_project/blob/main/LICENSE
[linkedin-shield]: https://img.shields.io/badge/-LinkedIn-black.svg?style=for-the-badge&logo=linkedin&colorB=555
[linkedin-url]: https://linkedin.com/in/naravid
[Django]: https://img.shields.io/badge/Django-092E20?style=for-the-badge&logo=django&logoColor=white
[Django-url]: https://www.djangoproject.com/
[TailwindCSS]: https://img.shields.io/badge/tailwindcss-%2338B2AC.svg?style=for-the-badge&logo=tailwind-css&logoColor=white
[TailwindCSS-url]: https://tailwindcss.com/
[Flowbite]: https://img.shields.io/badge/Flowbite-1C64F2?style=for-the-badge&logo=flowbite&logoColor=white
[Flowbite-url]: https://flowbite.com/
[Pandas]: https://img.shields.io/badge/pandas-%23150458.svg?style=for-the-badge&logo=pandas&logoColor=white
[Pandas-url]: https://pandas.pydata.org/
[JQuery.com]: https://img.shields.io/badge/jQuery-0769AD?style=for-the-badge&logo=jquery&logoColor=white
[JQuery-url]: https://jquery.com
