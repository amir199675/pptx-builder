# from pptx import Presentation
# from pptx.util import Inches, Pt
# import json
# import copy
#
# # شبیه‌سازی ورودی JSON
# json_data = [
#     {"template_slide": 0, "text": "سلام دنیا"},
#     {"template_slide": 1, "text": "سوال اول"},
#     {"template_slide": 1, "text": "سوال دوم"},
#     {"template_slide": 1, "text": "سوال سوم"},
#     {"template_slide": 1, "text": "سوال چهارم"},
#     {"template_slide": 2, "text": "نتیجه‌گیری"}
# ]
#
# # مسیر فایل قالب (template)
# source_pptx = "template_4.pptx"
# destination_pptx = "output_4.pptx"
#
# # باز کردن پاورپوینت مبدا
# src = Presentation(source_pptx)
#
# # ساخت پاورپوینت جدید
# dst = Presentation()
# dst.slides._sldIdLst.clear()  # حذف اسلاید پیش‌فرض اولیه
#
# # اندازه صفحه مطابق فایل مبدا
# dst.slide_width = src.slide_width
# dst.slide_height = src.slide_height
#
# # پردازش هر ورودی از JSON
# for entry in json_data:
#     src_slide_index = entry["template_slide"]
#     custom_text = entry["text"]
#
#     template_slide = src.slides[src_slide_index]
#
#     # ایجاد اسلاید جدید با layout خالی (برای کنترل کامل)
#     blank_layout = dst.slide_layouts[6]  # معمولاً "Blank" layout
#     new_slide = dst.slides.add_slide(blank_layout)
#
#     # کپی تمام اشکال (فقط متن و تصویر)
#     for shape in template_slide.shapes:
#         if shape.shape_type == 13:  # Picture
#             image = shape.image
#             with open("temp_img.png", "wb") as f:
#                 f.write(image.blob)
#             new_slide.shapes.add_picture("temp_img.png", shape.left, shape.top, shape.width, shape.height)
#
#         elif shape.has_text_frame:
#             textbox = new_slide.shapes.add_textbox(shape.left, shape.top, shape.width, shape.height)
#             tf = textbox.text_frame
#             for p in shape.text_frame.paragraphs:
#                 new_paragraph = tf.add_paragraph()
#                 new_paragraph.text = p.text
#                 new_paragraph.font.size = Pt(18)
#             tf.text = ""  # پاک‌کردن متن اول
#             tf.paragraphs[0].text = custom_text  # درج متن سفارشی
#
# # ذخیره فایل خروجی
# dst.save(destination_pptx)
# print(f"✅ فایل جدید ساخته شد: {destination_pptx}")


# import zipfile
# import os
# import shutil
# import tempfile
# import xml.etree.ElementTree as ET
# import json
#
# # ورودی JSON شبیه‌سازی‌شده
# json_data = [
#     {"template_slide": 1, "text": "سلام دنیا"},
#     {"template_slide": 2, "text": "مرحله دوم"},
#     {"template_slide": 3, "text": "نتیجه‌گیری"}
# ]
#
# def clone_slides_with_text(src_path, json_data):
#     temp_src = tempfile.mkdtemp()
#     temp_dst = tempfile.mkdtemp()
#
#     # استخراج فایل پاورپوینت مبدا
#     with zipfile.ZipFile(src_path, 'r') as zip_ref:
#         zip_ref.extractall(temp_src)
#
#     # کپی کل ساختار به مقصد
#     shutil.copytree(temp_src, temp_dst, dirs_exist_ok=True)
#
#     slide_folder = os.path.join(temp_dst, 'ppt', 'slides')
#     rels_folder = os.path.join(slide_folder, '_rels')
#
#     # تعداد اسلایدهای فعلی در فایل
#     existing_slides = [f for f in os.listdir(slide_folder) if f.startswith("slide") and f.endswith(".xml")]
#     base_slide_count = len(existing_slides)
#
#     # فایل‌های اصلی
#     presentation_xml = os.path.join(temp_dst, 'ppt', 'presentation.xml')
#     presentation_rels = os.path.join(temp_dst, 'ppt', '_rels', 'presentation.xml.rels')
#
#     pres_tree = ET.parse(presentation_xml)
#     pres_root = pres_tree.getroot()
#     sldIdLst = pres_root.find('{http://schemas.openxmlformats.org/presentationml/2006/main}sldIdLst')
#
#     rels_tree = ET.parse(presentation_rels)
#     rels_root = rels_tree.getroot()
#
#     # ایجاد اسلایدهای جدید بر اساس JSON
#     for i, item in enumerate(json_data):
#         original_index = item["template_slide"]
#         custom_text = item["text"]
#
#         old_slide = f"slide{original_index}.xml"
#         new_slide_index = base_slide_count + i + 1
#         new_slide = f"slide{new_slide_index}.xml"
#
#         # کپی فایل اسلاید
#         shutil.copy(os.path.join(temp_src, 'ppt', 'slides', old_slide),
#                     os.path.join(slide_folder, new_slide))
#
#         # کپی فایل rels
#         old_rels = os.path.join(temp_src, 'ppt', 'slides', '_rels', f"{old_slide}.rels")
#         new_rels = os.path.join(rels_folder, f"{new_slide}.rels")
#         if os.path.exists(old_rels):
#             shutil.copy(old_rels, new_rels)
#
#         # افزودن متن سفارشی در اسلاید جدید
#         slide_path = os.path.join(slide_folder, new_slide)
#         slide_tree = ET.parse(slide_path)
#         slide_root = slide_tree.getroot()
#
#         # اضافه کردن یک shape جدید با متن دلخواه
#         ns = {
#             'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
#             'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
#         }
#         spTree = slide_root.find('p:cSld/p:spTree', ns)
#         shape_id = 999 + i
#
#         new_shape = ET.fromstring(f'''
#         <p:sp xmlns:p="{ns['p']}" xmlns:a="{ns['a']}">
#           <p:nvSpPr>
#             <p:cNvPr id="{shape_id}" name="Custom Text"/>
#             <p:cNvSpPr/>
#             <p:nvPr/>
#           </p:nvSpPr>
#           <p:spPr/>
#           <p:txBody>
#             <a:bodyPr/>
#             <a:lstStyle/>
#             <a:p>
#               <a:r>
#                 <a:rPr lang="en-US" dirty="0" smtClean="0"/>
#                 <a:t>{custom_text}</a:t>
#               </a:r>
#               <a:endParaRPr lang="en-US" dirty="0"/>
#             </a:p>
#           </p:txBody>
#         </p:sp>
#         ''')
#         spTree.append(new_shape)
#         slide_tree.write(slide_path, xml_declaration=True, encoding='utf-8')
#
#         # افزودن به presentation.xml
#         new_id = 256 + new_slide_index
#         new_rid = f"rId{new_id}"
#         sldId = ET.Element('{http://schemas.openxmlformats.org/presentationml/2006/main}sldId')
#         sldId.set('id', str(new_id))
#         sldId.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', new_rid)
#         sldIdLst.append(sldId)
#
#         # افزودن به rels
#         new_rel = ET.Element('Relationship')
#         new_rel.set('Id', new_rid)
#         new_rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide')
#         new_rel.set('Target', f'slides/{new_slide}')
#         rels_root.append(new_rel)
#
#     # ذخیره تغییرات
#     pres_tree.write(presentation_xml, encoding='utf-8', xml_declaration=True)
#     rels_tree.write(presentation_rels, encoding='utf-8', xml_declaration=True)
#
#     # فشرده‌سازی به عنوان فایل جدید
#     out_pptx = src_path.replace('.pptx', '_generated.pptx')
#     with zipfile.ZipFile(out_pptx, 'w', zipfile.ZIP_DEFLATED) as pptx:
#         for folder_name, subfolders, filenames in os.walk(temp_dst):
#             for filename in filenames:
#                 abs_path = os.path.join(folder_name, filename)
#                 rel_path = os.path.relpath(abs_path, temp_dst)
#                 pptx.write(abs_path, rel_path)
#
#     shutil.rmtree(temp_src)
#     shutil.rmtree(temp_dst)
#     print(f"✅ فایل خروجی ساخته شد: {out_pptx}")
#
#
# # اجرای تابع
# clone_slides_with_text("template_4.pptx", json_data)


# from pptx import Presentation
# from pptx.util import Pt
# import json
#
# # ورودی JSON
# json_data = [
#     {"template_slide": 0, "text": "سلام دنیا"},
#     {"template_slide": 1, "text": "مرحله دوم"},
#     {"template_slide": 2, "text": "نتیجه‌گیری"}
# ]
#
# source_pptx = "template_4.pptx"
# destination_pptx = "output_4.pptx"
#
# # باز کردن پاورپوینت مبدا
# src = Presentation(source_pptx)
#
# # ساخت پاورپوینت جدید از روی source (با layoutهای مشابه)
# dst = Presentation()
# dst.slide_width = src.slide_width
# dst.slide_height = src.slide_height
# dst.slides._sldIdLst.clear()
#
# # کپی اسلاید با حفظ layout و جایگذاری متن سفارشی
# for entry in json_data:
#     index = entry["template_slide"]
#     text = entry["text"]
#
#     template_slide = src.slides[index]
#     layout = template_slide.slide_layout
#
#     # اضافه‌کردن اسلاید با همان layout
#     new_slide = dst.slides.add_slide(layout)
#
#     # کپی اشکال و جایگزینی متن
#     for shape in template_slide.shapes:
#         if shape.has_text_frame:
#             left = shape.left
#             top = shape.top
#             width = shape.width
#             height = shape.height
#             new_shape = new_slide.shapes.add_textbox(left, top, width, height)
#             new_tf = new_shape.text_frame
#             new_tf.text = text
#         elif shape.shape_type == 13:  # Picture
#             image = shape.image
#             with open("temp_img.png", "wb") as f:
#                 f.write(image.blob)
#             new_slide.shapes.add_picture("temp_img.png", shape.left, shape.top, shape.width, shape.height)
#
# # ذخیره فایل نهایی
# dst.save(destination_pptx)
# print(f"✅ ساخته شد: {destination_pptx}")



# import zipfile
# import os
# import shutil
# import tempfile
# import xml.etree.ElementTree as ET
#
# def clone_slide_and_replace_text(src_path, slide_index, custom_text):
#     ET.register_namespace('', "http://schemas.openxmlformats.org/presentationml/2006/main")
#     ET.register_namespace('a', "http://schemas.openxmlformats.org/drawingml/2006/main")
#     ET.register_namespace('r', "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
#
#     temp_src = tempfile.mkdtemp()
#     temp_dst = tempfile.mkdtemp()
#
#     # Unzip source pptx
#     with zipfile.ZipFile(src_path, 'r') as zip_ref:
#         zip_ref.extractall(temp_src)
#
#     # Copy everything to destination folder
#     shutil.copytree(temp_src, temp_dst, dirs_exist_ok=True)
#
#     # Keep only necessary folders
#     slide_id = slide_index + 1
#     all_slides = os.listdir(os.path.join(temp_dst, 'ppt', 'slides'))
#     all_slide_files = [f for f in all_slides if f.endswith('.xml')]
#     all_slide_rels = os.listdir(os.path.join(temp_dst, 'ppt', 'slides', '_rels'))
#
#     # Remove all other slides
#     for f in all_slide_files:
#         if f != f"slide{slide_id}.xml":
#             os.remove(os.path.join(temp_dst, 'ppt', 'slides', f))
#     for f in all_slide_rels:
#         if f != f"slide{slide_id}.xml.rels":
#             os.remove(os.path.join(temp_dst, 'ppt', 'slides', '_rels', f))
#
#     # Rename selected slide to slide1
#     os.rename(os.path.join(temp_dst, 'ppt', 'slides', f"slide{slide_id}.xml"),
#               os.path.join(temp_dst, 'ppt', 'slides', "slide1.xml"))
#     os.rename(os.path.join(temp_dst, 'ppt', 'slides', '_rels', f"slide{slide_id}.xml.rels"),
#               os.path.join(temp_dst, 'ppt', 'slides', '_rels', "slide1.xml.rels"))
#
#     # Update presentation.xml
#     presentation_xml = os.path.join(temp_dst, 'ppt', 'presentation.xml')
#     pres_tree = ET.parse(presentation_xml)
#     pres_root = pres_tree.getroot()
#     sldIdLst = pres_root.find('{http://schemas.openxmlformats.org/presentationml/2006/main}sldIdLst')
#     for sldId in list(sldIdLst):
#         sldIdLst.remove(sldId)
#     new_sldId = ET.SubElement(sldIdLst, '{http://schemas.openxmlformats.org/presentationml/2006/main}sldId')
#     new_sldId.set('id', '256')
#     new_sldId.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id', 'rId1')
#     pres_tree.write(presentation_xml, encoding='utf-8', xml_declaration=True)
#
#     # Update presentation.xml.rels
#     pres_rels = os.path.join(temp_dst, 'ppt', '_rels', 'presentation.xml.rels')
#     rels_tree = ET.parse(pres_rels)
#     rels_root = rels_tree.getroot()
#     for rel in list(rels_root):
#         if rel.attrib.get('Type', '').endswith('/slide'):
#             rels_root.remove(rel)
#     new_rel = ET.SubElement(rels_root, 'Relationship')
#     new_rel.set('Id', 'rId1')
#     new_rel.set('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide')
#     new_rel.set('Target', 'slides/slide1.xml')
#     rels_tree.write(pres_rels, encoding='utf-8', xml_declaration=True)
#
#     # Replace all text inside slide1.xml
#     slide_xml_path = os.path.join(temp_dst, 'ppt', 'slides', 'slide1.xml')
#     slide_tree = ET.parse(slide_xml_path)
#     slide_root = slide_tree.getroot()
#     for elem in slide_root.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}t'):
#         elem.text = custom_text
#     slide_tree.write(slide_xml_path, encoding='utf-8', xml_declaration=True)
#
#     # Zip to new pptx
#     output_path = src_path.replace('.pptx', f'_slide{slide_id}_custom.pptx')
#     with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as pptx:
#         for folder_name, subfolders, filenames in os.walk(temp_dst):
#             for filename in filenames:
#                 abs_path = os.path.join(folder_name, filename)
#                 rel_path = os.path.relpath(abs_path, temp_dst)
#                 pptx.write(abs_path, rel_path)
#
#     shutil.rmtree(temp_src)
#     shutil.rmtree(temp_dst)
#
#     print(f"✅ فایل ساخته شد: {output_path}")
#
#
# # مثال: کپی اسلاید دوم و جایگزینی متن با "سلام دنیا"
# clone_slide_and_replace_text("template_4.pptx", slide_index=1, custom_text="سلام دنیا")



# from pptx import Presentation
# from pptx.util import Pt
# from pptx.chart.data import CategoryChartData
# import json
#
# json_data = [
#     {
#         "template_slide": 0,
#         "text": "سلام دنیا"
#     },
#     {
#         "template_slide": 1,
#         "text": "سوال اول من",
#         "chart_data": {
#             "categories": ["الف", "ب", "ج"],
#             "series": {
#                 "درصد پاسخ صحیح": [80, 60, 90]
#             }
#         }
#     },
#     {
#         "template_slide": 2,
#         "text": "نتیجه‌گیری"
#     }
# ]
#
# source_pptx = "template_4.pptx"
# destination_pptx = "output_4.pptx"
#
# src = Presentation(source_pptx)
# dst = Presentation()
# dst.slides._sldIdLst.clear()
# dst.slide_width = src.slide_width
# dst.slide_height = src.slide_height
#
# for entry in json_data:
#     src_slide_index = entry["template_slide"]
#     custom_text = entry["text"]
#     chart_info = entry.get("chart_data")
#
#     template_slide = src.slides[src_slide_index]
#     layout = dst.slide_layouts[6]
#     new_slide = dst.slides.add_slide(layout)
#
#     for shape in template_slide.shapes:
#         if shape.shape_type == 13:  # Picture
#             image = shape.image
#             with open("temp_img.png", "wb") as f:
#                 f.write(image.blob)
#             new_slide.shapes.add_picture("temp_img.png", shape.left, shape.top, shape.width, shape.height)
#
#         elif shape.has_text_frame:
#             textbox = new_slide.shapes.add_textbox(shape.left, shape.top, shape.width, shape.height)
#             tf = textbox.text_frame
#             tf.text = custom_text
#             for p in shape.text_frame.paragraphs[1:]:
#                 tf.add_paragraph().text = p.text
#
#         elif shape.has_chart and chart_info:
#             chart = shape.chart
#             chart_data = CategoryChartData()
#             chart_data.categories = chart_info["categories"]
#             for series_name, values in chart_info["series"].items():
#                 chart_data.add_series(series_name, values)
#             chart.replace_data(chart_data)
#
# dst.save(destination_pptx)
# print(f"✅ ساخته شد: {destination_pptx}")

# from pptx import Presentation
# from pptx.util import Pt
# from pptx.chart.data import CategoryChartData
#
# json_data = [
#     {
#         "template_slide": 0,
#         "text": "سلام دنیا"
#     },
#     {
#         "template_slide": 1,
#         "text": "سوال اول",
#         "chart_data": {
#             "categories": ["الف", "ب", "ج"],
#             "series": {
#                 "درصد پاسخ صحیح": [80, 60, 90]
#             }
#         }
#     }
# ]
#
# source_pptx = "template_4.pptx"
# destination_pptx = "output_4.pptx"
#
# src = Presentation(source_pptx)
# dst = Presentation()
# dst.slides._sldIdLst.clear()
# dst.slide_width = src.slide_width
# dst.slide_height = src.slide_height
#
# for entry in json_data:
#     src_slide_index = entry["template_slide"]
#     custom_text = entry["text"]
#     chart_info = entry.get("chart_data")
#
#     # کپی layout همان اسلاید مبدا (نه layout خالی)
#     template_slide = src.slides[src_slide_index]
#     layout = template_slide.slide_layout
#     new_slide = dst.slides.add_slide(layout)
#
#     # کپی اشکال از مبدا به مقصد
#     for shape in template_slide.shapes:
#         if shape.has_text_frame:
#             textbox = new_slide.shapes.add_textbox(shape.left, shape.top, shape.width, shape.height)
#             tf = textbox.text_frame
#             tf.text = custom_text
#
#         elif shape.has_chart and chart_info:
#             # انتقال چارت به اسلاید جدید
#             chart = shape.chart
#             chart_data = CategoryChartData()
#             chart_data.categories = chart_info["categories"]
#             for series_name, values in chart_info["series"].items():
#                 chart_data.add_series(series_name, values)
#             chart.replace_data(chart_data)
#
#         elif shape.shape_type == 13:  # Picture
#             image = shape.image
#             with open("temp_img.png", "wb") as f:
#                 f.write(image.blob)
#             new_slide.shapes.add_picture("temp_img.png", shape.left, shape.top, shape.width, shape.height)
#
# dst.save(destination_pptx)
# print(f"✅ فایل خروجی ساخته شد: {destination_pptx}")
