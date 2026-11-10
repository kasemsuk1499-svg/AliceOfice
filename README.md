# AliceOfice — KNS Logistics Workspace

เว็บสำหรับเตรียมโหลด STP/PTR และซิงก์ Loading → Export โดยใช้ Excel แบบฟอร์มเดิม พร้อม Supabase สำหรับเก็บไฟล์ต้นแบบและงานย้อนหลังแบบส่วนตัว

## วิธีใช้

1. เปิด [Alice Office](https://kasemsuk1499-svg.github.io/AliceOfice/) หลังเปิด GitHub Pages ตามขั้นตอนด้านล่าง
2. เข้าสู่ระบบ Alice Cloud ทางอีเมล เพื่อเปิดประวัติและการซิงก์หลายอุปกรณ์
3. เลือกวันที่และงาน **เตรียมโหลดวันถัดไป** หรือ **Daily Report**
4. อัปโหลดไฟล์ต้นทาง หรือเลือกแม่แบบล่าสุดที่บันทึกบนคลาวด์ → กด **ตรวจสอบไฟล์และสร้างรายการ**
5. ตรวจ Booking / รถ / ตู้ / น้ำหนักก่อนกด **สร้างไฟล์ Excel**, ดาวน์โหลดและตรวจ Print Preview

**เตรียมโหลด:** รับ CSV STP1/STP2 และแม่แบบ STP/PTR XLSX. กรองเฉพาะ `Load = KNS` และตัด Cancel/แคนเซิล/ยกเลิก. ตั้งชื่อชีตตาม Booking, สองตู้ต่อชีต. STP ยก `Packaging and Batch no.` ทั้งข้อความลง Text Box. PTR คง Gate No. และใส่ชื่อแป้งกับจำนวนถุง. 20 ถุง = Pallet, จำนวนอื่น = แบก. ตารางเขียนมือคงว่าง.

**แผนแบช Rev (.xlsx):** อัปโหลดไฟล์ `Rev 1 แผนส่งออกวันที่ 10 พ.ย. 2569.xlsx` แล้วเลือกประเภท `แผนแบช / ตำแหน่งจัดเก็บ (Rev)` (ระบบจำแนกชื่อไฟล์ให้อัตโนมัติ). กดตรวจสอบเพื่อเปิดหน้ารายละเอียด 8 ชีต รวมสต็อกอ้างอิงใน `Sheet1` และรายการ Qty Allocate ในชีตสินค้า. แยก `Batch_Code`, Qty, UOM, Inventory Status และตัวเลขจัดสรร (ค่า Location ยังคงอยู่ในไฟล์ต้นทาง แต่ไม่ได้แสดงในตารางหลัก) เช่น `1(6) 2(12)`. กรองตามชีต คัน/ตู้ หรือค้นหา Batch/Code/สินค้า ได้. เมื่อใช้ร่วมกับ DAILY จะเทียบ Batch + Code + ลำดับคันกับ Loading เท่าที่จับคู่แน่ชัดได้และแจ้งความต่าง; ไม่เขียนทับยอดโหลดจริง ไม่สร้างรถหรือแบชจากแผนโดยอัตโนมัติ. แผนที่จัดสรรหลาย UOM จะไม่ถูกรวมเป็นยอดถุงเดียวกันโดยไม่แยกหน่วย.

**Daily:** อัปโหลด Loading STP/PTR, Domestic / SSAMS Loading, Export STP/PTR และไฟล์ขนส่ง (ถ้ามี). กติกา: ทุก Loading ที่ไม่ใช่ PTR (รวม Domestic / SSAMS) ลง Export STP ร่วมกันในไฟล์เดียว; Loading PTR ลง Export PTR เท่านั้น. Daily ใช้รายการจริงใน Loading ทั้งหมด แม้ไม่ได้อยู่ใน CSV แผน. จับคู่แถวเดิมด้วย Booking+Container หรือ Weight Card/ทะเบียนรถสำหรับ Domestic ซึ่งไม่มี Booking/Container; ไม่ใช้เลขคันใน Loading ไปแทนลำดับรวม Export, ไม่เพิ่มรายการซ้ำ, และรักษาประวัติ. นำข้อมูลทุกช่องที่มีคู่ใน Export ไปลง; ช่องไม่มีข้อมูลจริงยังคงว่าง.

**กรอก Export เพิ่มเติมบนเว็บ:** ใน DAILY ให้คลิก “กรอกข้อมูล Export เพิ่มเติม” ใต้รถแต่ละคัน เพื่อกรอก Invoice, Contract, Agent, Place of Receipt/Return, Truck Company, Driver, Price, Overnight Charge, Time Arrive/Out, Root Cause Delay และคะแนนประเมิน. ค่าเริ่มต้นที่พบซ้ำจริงในไฟล์รายงานเก่า เช่น `-`, เวลาโหลด `10:00`–`10:50`, คะแนนเอกสาร/ขนส่ง `1` เป็นค่าเริ่มต้นแก้ไขได้ ไม่ใช่หลักฐานเวลารถเข้าจริง. ระบบคำนวณ Total Time และคะแนนรวมจากค่าที่แก้ และเขียนลงคอลัมน์ STP/PTR ตามฟอร์มเดิม. เมื่อเข้าสู่ระบบ Alice Cloud ข้อมูลที่กรอกจะถูกบันทึกเป็นฉบับร่างและเรียกกลับมาแก้ได้.

**Cloud:** หลังเข้าสู่ระบบ เมื่อสร้างไฟล์สำเร็จ ระบบเก็บต้นทาง ผลลัพธ์ และรายการรถให้โดยอัตโนมัติ; เรียกแม่แบบล่าสุดหรือเปิดงานย้อนหลังจากอีกเครื่องได้. ซิงก์ Realtime แจ้งเมื่อรายการเปลี่ยน โดยไม่เขียนทับงานที่กำลังแก้เงียบ ๆ.

## เปิดใช้งานเว็บไซต์ — ต้องตั้งค่าครั้งเดียว

**GitHub:** ไป [Settings → Pages](https://github.com/kasemsuk1499-svg/AliceOfice/settings/pages), เลือก **Build and deployment → Source: GitHub Actions** แล้วรันซ้ำ workflow [Deploy Alice Office to GitHub Pages](https://github.com/kasemsuk1499-svg/AliceOfice/actions). มี workflow อยู่ที่ `.github/workflows/deploy-pages.yml` แล้ว.

**Supabase:** ในโปรเจกต์ `qlaykelpabbjojpqjfwi`, ไป **Authentication → URL Configuration → Redirect URLs** และเพิ่มลิงก์นี้ให้ Magic Link กลับมายังเว็บ:

```text
https://kasemsuk1499-svg.github.io/AliceOfice/
```

**อย่าเปลี่ยน Site URL เดิม** ของ Supabase เพราะโปรเจกต์นี้อาจมีบริการอื่นใช้อยู่. ให้เพิ่มแค่ Redirect URL ของ Alice Office.

## ความปลอดภัยและความต่อเนื่อง

- ตาราง `alice_office_drafts`, `alice_office_files`, `alice_office_vehicles` แยกจากตารางอื่น และใช้ RLS จำกัดเจ้าของข้อมูล (`auth.uid() = owner_id`)
- Bucket `alice-office-private` เป็น Private; วางไฟล์ภายใต้โฟลเดอร์ของ user ID และกำหนด Storage RLS
- ไม่เปิดข้อมูล Loading/Export บน GitHub. ใน repository มีเฉพาะโค้ดและ Supabase publishable key (ไม่ใช่ secret key)
- ผู้ใช้ต้องล็อกอินบัญชีเดิมบนทุกอุปกรณ์เพื่อดูประวัติต่อเนื่อง; ยังไม่มีแชร์งานหลายบัญชี
- สามารถกด **เก็บแม่แบบล่าสุด**, **ใช้ไฟล์** หรือ **เปิดงานย้อนหลัง**; ผลลัพธ์ Export/Loading ที่สร้างจะถูกเก็บเป็นแม่แบบเวอร์ชันล่าสุด
- Migration อยู่ใน `supabase/migrations/20261109_alice_office.sql`

## สถานะทดสอบและข้อจำกัด

- เตรียมโหลดวันที่ 10 พ.ย. 2026: STP 20 / PTR 6 = 26 ตู้, เปิดไฟล์ Excel ได้
- Daily วันที่ 9 พ.ย. 2026: STP 18 / PTR 1 = 19 คันตรงแถวเดิม, เพิ่มแถวซ้ำ **0**, Excel เปิดได้หลังแก้ calcChain
- Daily วันที่ 10 พ.ย. 2026: ทดสอบ Domestic / SSAMS จริงผ่านหน้าเว็บ → Export STP ได้ 1 คัน, LT-20 38 ถุง/32,300 กก., และรันซ้ำไม่เกิดแถวซ้ำ
- Daily ต้องมี Loading ที่มีข้อมูลจริงอยู่ก่อน; การกระจาย Batch/Tank Code เพื่อสร้าง Loading จากแผนล้วน ๆ ยังไม่ทำอัตโนมัติเพราะห้ามเดา
- งานพิเศษ REKKIE และ C*Cream ยังรอคู่มือเฉพาะ
- การยืนยันว่า Cloud เชื่อมตั้งแต่ล็อกอินจนบันทึกไฟล์ได้จริง ต้องทดสอบด้วยบัญชีผู้ใช้หลังเปิด GitHub Pages และ Redirect URL

Stack: static HTML + JSZip + SheetJS + Supabase Auth/Storage/Realtime.