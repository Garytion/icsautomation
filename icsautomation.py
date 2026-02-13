import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import zipfile
import os
import shutil
import tempfile
import io

# 页面配置
st.set_page_config(page_title="Excel 业务自动化整合工具", layout="wide")

st.title("📂 ICS2业务自动化整合工具")
st.markdown("""
### 使用说明：
1. **上传 `containerinformation.xlsx`**: 包含单号、件数、重量、HS CODE等原始数据。
2. **上传 `icstemplate.xlsx`**: 作为生成单号文件的模板。
3. **上传 `realdoc.zip`**: 包含额外补充信息的压缩包，文件名应与单号一致。
4. **点击处理**: 程序将自动生成中间文件并进行二次填充，最后提供下载。
""")

def process_logic():
    # 界面上传
    col1, col2, col3 = st.columns(3)
    with col1:
        container_file = st.file_uploader("1. 上传 containerinformation", type=["xlsx"])
    with col2:
        template_file = st.file_uploader("2. 上传 icstemplate", type=["xlsx"])
    with col3:
        realdoc_zip = st.file_uploader("3. 上传 realdoc.zip", type=["zip"])

    if st.button("🔥 开始执行全流程处理"):
        if not (container_file and template_file and realdoc_zip):
            st.error("请确保三个文件都已上传！")
            return

        with st.spinner("正在深度处理中，请稍候..."):
            try:
                # 使用临时目录处理文件
                with tempfile.TemporaryDirectory() as tmp_dir:
                    r_dir = os.path.join(tmp_dir, "R")
                    p_dir = os.path.join(tmp_dir, "P")
                    out_dir = os.path.join(tmp_dir, "Output")
                    os.makedirs(r_dir)
                    os.makedirs(p_dir)
                    os.makedirs(out_dir)

                    # --- 步骤 1: 处理需求一 (生成中间文件 P) ---
                    df = pd.read_excel(container_file)
                    df['单号'] = df['单号'].ffill() # 自动填充空单号
                    
                    template_bytes = template_file.read()
                    grouped = df.groupby('单号')

                    for bill_no, group in grouped:
                        # 加载模板
                        wb = load_workbook(io.BytesIO(template_bytes))
                        ws = wb.active

                        # 填充头部汇总信息
                        ws['B5'] = bill_no
                        ws['B8'] = group['件数'].sum()
                        ws['B9'] = group['重量(KGS)'].sum()

                        # 获取 F130 的原始值用于向下填充
                        f130_initial_val = ws['F130'].value

                        # 填充明细行 (从 130 行开始)
                        for i, (_, row) in enumerate(group.iterrows()):
                            current_row = 130 + i
                            ws[f'A{current_row}'] = row['HS CODE']
                            ws[f'B{current_row}'] = row['品名']
                            ws[f'C{current_row}'] = row['件数']
                            ws[f'D{current_row}'] = "PK-Package"
                            ws[f'E{current_row}'] = row['重量(KGS)']
                            ws[f'F{current_row}'] = f130_initial_val
                        
                        # 保存到临时 P 文件夹
                        wb.save(os.path.join(p_dir, f"{bill_no}.xlsx"))

                    # --- 步骤 2: 处理需求二 (从 realdoc 提取数据到 P) ---
                    # 解压 realdoc 到 R
                    with zipfile.ZipFile(realdoc_zip, 'r') as z:
                        z.extractall(r_dir)

                    # 定义映射关系
                    row_mapping = {7: 14, 8: 15, 10: 18, 11: 19}
                    columns = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']

                    # 匹配并合并
                    p_files = [f for f in os.listdir(p_dir) if f.endswith('.xlsx')]
                    match_count = 0
                    
                    for filename in p_files:
                        p_path = os.path.join(p_dir, filename)
                        r_path = os.path.join(r_dir, filename)

                        if os.path.exists(r_path):
                            wb_p = load_workbook(p_path)
                            ws_p = wb_p.active
                            wb_r = load_workbook(r_path, data_only=True)
                            ws_r = wb_r.active

                            for src_row, tgt_row in row_mapping.items():
                                for col in columns:
                                    ws_p[f"{col}{tgt_row}"].value = ws_r[f"{col}{src_row}"].value
                            
                            wb_p.save(os.path.join(out_dir, filename))
                            match_count += 1
                        else:
                            # 如果 realdoc 里没找到，也将中间件移过去，方便排查
                            shutil.copy(p_path, os.path.join(out_dir, filename))

                    # --- 步骤 3: 打包最终结果 ---
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w") as z:
                        for f in os.listdir(out_dir):
                            z.write(os.path.join(out_dir, f), arcname=f)
                    
                    st.success(f"✅ 处理完成！共处理 {len(p_files)} 个单号，其中 {match_count} 个成功匹配 realdoc 补充信息。")
                    st.download_button(
                        label="📥 点击下载最终结果包 (.zip)",
                        data=zip_buffer.getvalue(),
                        file_name="final_results.zip",
                        mime="application/zip"
                    )

            except Exception as e:
                st.error(f"发生错误: {e}")

if __name__ == "__main__":
    process_logic()