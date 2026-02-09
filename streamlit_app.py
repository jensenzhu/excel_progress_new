import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl import load_workbook
import tempfile
import os
import difflib

st.set_page_config(
    page_title="Excel数据处理工具",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Excel数据处理工具")
st.markdown("---")

st.markdown("### 📁 文件上传")

col1, col2 = st.columns(2)

with col1:
    st.markdown("#### ERP库存表（from文件）")
    from_file = st.file_uploader(
        "上传ERP库存表",
        type=['xlsx', 'xls', 'csv'],
        key='from_file',
        help="上传包含库存数据的Excel文件"
    )

with col2:
    st.markdown("#### 订单表（dist文件）")
    dist_file = st.file_uploader(
        "上传订单表",
        type=['xlsx', 'xls'],
        key='dist_file',
        help="上传需要更新的订单Excel文件"
    )

st.markdown("---")

st.markdown("### ⚙️ 处理配置")

target_column = st.text_input(
    "要填入的列名称",
    placeholder="例如：所需数量/个（标箱倍数）",
    help="输入目标Excel文件中要更新数据的列名称"
)

st.markdown("---")

st.markdown("### 🚀 开始处理")

if st.button("开始处理", type="primary", use_container_width=True):
    if not from_file:
        st.error("❌ 请先上传ERP库存表（from文件）")
        st.stop()
    
    if not dist_file:
        st.error("❌ 请先上传订单表（dist文件）")
        st.stop()
    
    if not target_column.strip():
        st.error("❌ 请输入要填入的列名称")
        st.stop()
    
    with st.spinner("正在处理数据..."):
        try:
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            status_text.text("📖 读取ERP库存表...")
            progress_bar.progress(10)
            
            try:
                df_source = pd.read_excel(from_file, header=1)
                status_text.text(f"✅ 成功读取ERP库存表，共 {len(df_source)} 行数据")
            except Exception as e:
                st.error(f"❌ 读取ERP库存表失败: {str(e)}")
                st.stop()
            
            progress_bar.progress(30)
            
            status_text.text("🔍 提取产品型号...")
            
            merchant_code_cols = [col for col in df_source.columns if '商家' in col and '编码' in col]
            
            if not merchant_code_cols:
                st.error("❌ 未在ERP库存表中找到商家编码列")
                st.stop()
            
            def extract_model(code):
                if isinstance(code, str) and '-' in code:
                    parts = code.split('-')
                    if len(parts) >= 2:
                        return '-'.join(parts[1:])
                return None
            
            for col in merchant_code_cols:
                df_source[f'产品型号_{col}'] = df_source[col].apply(extract_model)
            
            model_cols = [col for col in df_source.columns if '产品型号_' in col]
            
            if not model_cols:
                st.error("❌ 未能提取任何产品型号")
                st.stop()
            
            df_source['产品型号'] = df_source[model_cols[0]]
            for col in model_cols[1:]:
                df_source['产品型号'] = df_source['产品型号'].fillna(df_source[col])
            
            erp_models = set(df_source['产品型号'].dropna().unique())
            status_text.text(f"✅ 成功提取产品型号，共 {len(erp_models)} 个")
            progress_bar.progress(50)
            
            status_text.text("📊 计算差值...")
            
            if '实际可用数' not in df_source.columns or '30天销量' not in df_source.columns:
                st.error("❌ ERP库存表中缺少'实际可用数'或'30天销量'列")
                st.stop()
            
            df_source['差值'] = df_source['30天销量'] - df_source['实际可用数']
            status_text.text(f"✅ 成功计算差值")
            progress_bar.progress(60)
            
            status_text.text("📖 读取订单表...")
            
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_dist:
                    tmp_dist.write(dist_file.read())
                    tmp_dist_path = tmp_dist.name
                
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_output:
                    tmp_output_path = tmp_output.name
                
                import shutil
                shutil.copy2(tmp_dist_path, tmp_output_path)
                
                wb = load_workbook(tmp_output_path, data_only=False, keep_links=True)
                ws = wb.active
                
                status_text.text(f"✅ 成功读取订单表，工作表名称: {ws.title}")
            except Exception as e:
                st.error(f"❌ 读取订单表失败: {str(e)}")
                st.stop()
            
            progress_bar.progress(70)
            
            status_text.text("🔍 查找目标列...")
            
            target_col_idx = None
            product_model_col_idx = None
            
            for row_idx in range(1, 6):
                for col_idx in range(1, ws.max_column + 1):
                    cell_value = ws.cell(row=row_idx, column=col_idx).value
                    if isinstance(cell_value, str):
                        if target_column in cell_value and not target_col_idx:
                            target_col_idx = col_idx
                            st.info(f"📍 在第 {row_idx} 行找到目标列 '{target_column}'，列索引: {target_col_idx}")
                        elif '产品型号' in cell_value and not product_model_col_idx:
                            product_model_col_idx = col_idx
                            st.info(f"📍 在第 {row_idx} 行找到产品型号列，列索引: {product_model_col_idx}")
                
                if target_col_idx and product_model_col_idx:
                    break
            
            if not target_col_idx:
                st.error(f"❌ 未在订单表中找到列 '{target_column}'，请检查列名称是否正确")
                st.stop()
            
            if not product_model_col_idx:
                st.error("❌ 未在订单表中找到产品型号列")
                st.stop()
            
            progress_bar.progress(80)
            
            status_text.text("🔄 更新数据...")
            
            model_diff_map = df_source.set_index('产品型号')['差值'].to_dict()
            
            order_models = set()
            updated_count = 0
            skipped_count = 0
            for row in range(4, ws.max_row + 1):
                model = ws.cell(row=row, column=product_model_col_idx).value
                
                if model:
                    order_models.add(model)
                    if model in model_diff_map:
                        diff_value = model_diff_map[model]
                        if diff_value >= 0:
                            ws.cell(row=row, column=target_col_idx).value = diff_value
                            updated_count += 1
                        else:
                            skipped_count += 1
            
            status_text.text(f"✅ 数据更新完成，共更新了 {updated_count} 个单元格，跳过 {skipped_count} 个负数")
            progress_bar.progress(90)
            
            status_text.text("💾 保存文件...")
            wb.save(tmp_output_path)
            
            with open(tmp_output_path, 'rb') as f:
                st.session_state['output_file'] = f.read()
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            st.session_state['output_filename'] = f"订单表_更新_{timestamp}.xlsx"
            
            progress_bar.progress(100)
            status_text.text("✅ 处理完成！")
            
            st.success(f"🎉 处理成功！共更新了 {updated_count} 个产品型号，跳过 {skipped_count} 个负数")
            
            os.unlink(tmp_dist_path)
            os.unlink(tmp_output_path)
            
            models_in_erp_not_in_order = sorted(erp_models - order_models)
            
            if models_in_erp_not_in_order:
                st.markdown("---")
                st.markdown("### ⚠️ ERP库存表中有但订单表中没有的产品型号")
                st.info(f"共找到 {len(models_in_erp_not_in_order)} 个产品型号在ERP库存表中存在，但在订单表中不存在：")
                
                def find_similar_model(target_model, all_models, threshold=0.6):
                    best_match = None
                    best_ratio = 0
                    for model in all_models:
                        ratio = difflib.SequenceMatcher(None, target_model, model).ratio()
                        if ratio >= threshold and ratio > best_ratio:
                            best_ratio = ratio
                            best_match = model
                    return best_match, best_ratio
                
                cols_per_row = 5
                for i in range(0, len(models_in_erp_not_in_order), cols_per_row):
                    cols = st.columns(cols_per_row)
                    for j, col in enumerate(cols):
                        if i + j < len(models_in_erp_not_in_order):
                            missing_model = models_in_erp_not_in_order[i + j]
                            similar_model, similarity = find_similar_model(missing_model, order_models)
                            
                            if similar_model:
                                col.markdown(f"**{missing_model}** → {similar_model} ({similarity*100:.0f}%)")
                            else:
                                col.markdown(f"**{missing_model}**")
            
        except Exception as e:
            st.error(f"❌ 处理过程中发生错误: {str(e)}")
            st.exception(e)
            st.stop()

st.markdown("---")

st.markdown("### 📥 下载结果")

if 'output_file' in st.session_state:
    st.download_button(
        label="📥 下载处理后的Excel文件",
        data=st.session_state['output_file'],
        file_name=st.session_state['output_filename'],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )
    st.info(f"📄 文件名: {st.session_state['output_filename']}")
else:
    st.info("💡 请先上传文件并点击'开始处理'按钮")

st.markdown("---")

st.markdown("### 📋 使用说明")
st.markdown("""
1. **上传ERP库存表**：上传包含库存数据的Excel文件（支持.xlsx, .xls, .csv格式）
2. **上传订单表**：上传需要更新的订单Excel文件（支持.xlsx, .xls格式）
3. **输入列名称**：输入订单表中要更新数据的列名称（例如：所需数量/个（标箱倍数））
4. **开始处理**：点击按钮开始处理数据
5. **下载结果**：处理完成后，点击下载按钮获取更新后的Excel文件

**注意事项：**
- ERP库存表需要包含"实际可用数"和"30天销量"列
- 订单表需要包含"产品型号"列和指定的目标列
- 系统会自动提取产品型号并计算差值（30天销量 - 实际可用数）
- 只有非负数的差值才会填入订单表，负数会被跳过
- 处理后的文件会保留原始格式和图片
- 会显示ERP库存表中有但订单表中没有的产品型号
- 对于缺失的型号，会显示订单表中相似度最高的型号（相似度≥60%）
""")