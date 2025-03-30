import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
from datetime import datetime
import os
from pathlib import Path
import json

def load_config(config_file='config.json'):
    """从配置文件加载设置"""
    try:
        if os.path.exists(config_file):
            with open(config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
            return config.get('folders', [])
        return []
    except Exception as e:
        st.error(f"读取配置文件出错：{str(e)}")
        return []

def save_config(folders, config_file='config.json'):
    """保存设置到配置文件"""
    try:
        config = {'folders': folders}
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)
        return True
    except Exception as e:
        st.error(f"保存配置文件出错：{str(e)}")
        return False

def get_all_word_files(folder_path):
    """获取文件夹下所有的Word文件"""
    word_files = []
    for file in Path(folder_path).rglob("*.docx"):
        word_files.append(file)
    return word_files

def extract_tables_from_word(file_path):
    """从Word文件中提取所有表格"""
    try:
        doc = Document(file_path)
        all_tables = []
        
        for table in doc.tables:
            # 提取表格数据
            table_data = []
            for row in table.rows:
                row_data = []
                for cell in row.cells:
                    row_data.append(cell.text.strip())
                table_data.append(row_data)
            
            # 转换为DataFrame
            if table_data:
                df = pd.DataFrame(table_data[1:], columns=table_data[0])
                all_tables.append(df)
        
        return all_tables
    except Exception as e:
        st.error(f"无法打开文件 {file_path.name}：{str(e)}")
        return []

def process_date_warnings(df, date_rules):
    """处理多组日期预警"""
    warning_results = []
    
    for rule in date_rules:
        start_col = rule['start_column']
        end_col = rule['end_column']
        warning_days = rule['warning_days']
        stability_days = rule['stability_days']
        
        if start_col not in df.columns:
            continue
            
        # 转换起始日期
        df[start_col] = pd.to_datetime(df[start_col], errors='coerce')
        
        # 获取终止日期
        if end_col and end_col in df.columns:
            end_dates = pd.to_datetime(df[end_col], errors='coerce')
        else:
            end_dates = pd.Timestamp.now()
            
        # 计算日期差
        days_diff = (end_dates - df[start_col]).dt.days
        
        # 计算剩余稳定性期限
        remaining_days = stability_days - days_diff
        
        # 标记超过预警天数的样品
        warning_mask = days_diff > warning_days
        
        if warning_mask.any():
            warning_results.append({
                'mask': warning_mask,
                'start_col': start_col,
                'end_col': end_col or '当前日期',
                'warning_days': warning_days,
                'stability_days': stability_days,
                'days_diff': days_diff,
                'remaining_days': remaining_days
            })
    
    return df, warning_results

def config_editor():
    """配置管理界面"""
    st.title("配置管理")
    st.markdown("---")

    # 使用tabs组织配置管理界面
    tab1, tab2 = st.tabs(["配置编辑", "导入导出"])
    
    with tab1:
        # 显示当前配置
        with st.expander("当前配置", expanded=True):
            current_config = {'folders': st.session_state.folders}
            st.json(current_config)
        
        # 添加新配置
        with st.form("add_config", clear_on_submit=True):
            st.subheader("添加新配置")
            
            col1, col2 = st.columns(2)
            with col1:
                folder_name = st.text_input("文件夹名称")
                folder_path = st.text_input("文件夹路径")
            with col2:
                start_col = st.text_input("起始日期列名称")
                end_col = st.text_input("终止日期列名称（可选）")
            
            col3, col4 = st.columns(2)
            with col3:
                warning_days = st.number_input("预警天数", value=180, min_value=1)
            with col4:
                stability_days = st.number_input("稳定性期限(天)", value=365, min_value=1)
            
            submit = st.form_submit_button("添加配置", use_container_width=True)
            
            if submit:
                new_config = {
                    'name': folder_name,
                    'path': folder_path,
                    'rule': {
                        'start_column': start_col,
                        'end_column': end_col if end_col else None,
                        'warning_days': warning_days,
                        'stability_days': stability_days
                    }
                }
                st.session_state.folders.append(new_config)
                save_config(st.session_state.folders)
                st.success("配置已添加")
                st.rerun()

    with tab2:
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("导出配置")
            if st.session_state.folders:
                config_data = {'folders': st.session_state.folders}
                config_json = json.dumps(config_data, ensure_ascii=False, indent=4)
                
                export_filename = st.text_input(
                    "导出文件名", 
                    value="config.json",
                    help="输入要保存的文件名（包括.json后缀）"
                )
                
                st.download_button(
                    label="导出配置文件",
                    data=config_json.encode('utf-8'),
                    file_name=export_filename,
                    mime="application/json",
                    help="下载配置文件到本地"
                )
        
        with col2:
            st.subheader("导入配置")
            uploaded_file = st.file_uploader(
                "导入配置文件", 
                type=['json'],
                help="选择要导入的配置文件（json格式）"
            )
            if uploaded_file is not None:
                try:
                    imported_config = json.load(uploaded_file)
                    if 'folders' in imported_config:
                        st.session_state.folders = imported_config['folders']
                        save_config(st.session_state.folders)
                        st.success("配置导入成功！")
                        st.rerun()
                    else:
                        st.error("无效的配置文件格式！")
                except Exception as e:
                    st.error(f"导入配置文件失败：{str(e)}")

def process_tables():
    """表格处理界面"""
    with st.container():
        st.title("样品预警系统")
        st.markdown("---")
    
    if 'merged_tables' not in st.session_state:
        st.session_state.merged_tables = {}
    if 'warning_data' not in st.session_state:
        st.session_state.warning_data = {}
    
    st.session_state.warning_data = {}
    
    with st.sidebar:
        st.title("设置")
        
        if 'folders' not in st.session_state:
            st.session_state.folders = load_config()
        
        st.subheader("添加新的文件夹")
        with st.form("add_folder"):
            folder_name = st.text_input("文件夹名称", key="new_folder_name")
            folder_path = st.text_input("文件夹路径", key="new_folder_path")
            
            st.write("预警规则设置：")
            start_col = st.text_input("起始日期列名称", key="new_start")
            end_col = st.text_input("终止日期列名称（可选）", key="new_end")
            warning_days = st.number_input("预警天数", value=180, min_value=1, key="new_days")
            stability_days = st.number_input("稳定性期限(天)", value=365, min_value=1, key="new_stability")
            
            if st.form_submit_button("添加"):
                new_folder = {
                    'name': folder_name,
                    'path': folder_path,
                    'rule': {
                        'start_column': start_col,
                        'end_column': end_col if end_col else None,
                        'warning_days': warning_days,
                        'stability_days': stability_days
                    }
                }
                st.session_state.folders.append(new_folder)
                save_config(st.session_state.folders)
        
        st.markdown("---")
        st.subheader("现有文件夹")
        
        for i, folder in enumerate(st.session_state.folders):
            st.markdown(f"""
            #### 📁 {folder['name'] or f'文件夹 {i+1}'}
            - 路径：`{folder['path']}`
            - 起始日期列：`{folder['rule']['start_column']}`
            - 终止日期列：`{folder['rule']['end_column'] or '当前日期'}`
            - 预警天数：{folder['rule']['warning_days']}天
            - 稳定性期限：{folder['rule']['stability_days']}天
            """)
            
            if st.button(f"删除文件夹 {i+1}", key=f"delete_{i}"):
                st.session_state.folders.pop(i)
                save_config(st.session_state.folders)
                st.rerun()
            st.markdown("---")
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("清除现有配置", type="secondary"):
                st.session_state.folders = []
                st.rerun()
        with col2:
            if st.button("清除所有配置", type="secondary"):
                st.session_state.folders = []
                save_config([])
                st.rerun()

    all_tables = []
    all_warning_samples = []
    merged_all = None
    combined_warnings = None
    status_summary = pd.Series()
    excel_buffer = None
    warning_buffer = None
    current_time = datetime.now().strftime("%Y%m%d_%H%M%S")

    folder_status = {
        'success': [],
        'empty': [],
        'error': [],
        'file_errors': {}
    }

    for folder_config in st.session_state.folders:
        folder_name = folder_config['name']
        folder_path = folder_config['path']
        
        try:
            if not os.path.exists(folder_path):
                folder_status['error'].append((folder_name, f"文件夹路径不存在: {folder_path}"))
                continue

            word_files = get_all_word_files(folder_path)
            
            if not word_files:
                folder_status['empty'].append(folder_name)
                continue
            
            folder_tables = []
            processed_files = 0
            failed_files = []
            
            for file_path in word_files:
                try:
                    tables = extract_tables_from_word(file_path)
                    if tables:
                        for table in tables:
                            table['文件名'] = file_path.name
                            table['来源文件夹'] = folder_name
                            folder_tables.extend([table])
                        processed_files += 1
                    else:
                        failed_files.append(file_path.name)
                except Exception as e:
                    failed_files.append(file_path.name)
                    continue
            
            if failed_files:
                folder_status['file_errors'][folder_name] = failed_files

            if folder_tables:
                folder_merged = pd.concat(folder_tables, ignore_index=True)
                all_tables.append(folder_merged)
                
                folder_merged, warning_results = process_date_warnings(folder_merged, [folder_config['rule']])
                
                date_col = folder_config['rule']['start_column']
                if date_col in folder_merged.columns:
                    folder_merged[date_col] = pd.to_datetime(folder_merged[date_col]).dt.strftime('%Y-%m-%d')
                
                end_col = folder_config['rule']['end_column']
                if end_col and end_col in folder_merged.columns:
                    folder_merged[end_col] = pd.to_datetime(folder_merged[end_col]).dt.strftime('%Y-%m-%d')
                
                for warning in warning_results:
                    warning_samples = folder_merged[warning['mask']].copy()
                    warning_samples['已用天数'] = warning['days_diff'][warning['mask']]
                    warning_samples['剩余稳定性期限(天)'] = warning['remaining_days'][warning['mask']]
                    
                    warning_samples['状态'] = warning_samples['剩余稳定性期限(天)'].apply(
                        lambda x: '❌ 已超期' if x <= 0 else ('⚠️ 接近超期' if x <= 30 else '✅ 正常')
                    )
                    warning_samples = warning_samples[warning_samples['状态'] != '✅ 正常']
                    
                    if not warning_samples.empty:
                        all_warning_samples.append(warning_samples)

                folder_status['success'].append((folder_name, processed_files, len(failed_files)))

        except Exception as e:
            folder_status['error'].append((folder_name, str(e)))

    st.header("处理状态摘要")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("成功处理的文件夹", len(folder_status['success']))
    with col2:
        st.metric("空文件夹", len(folder_status['empty']))
    with col3:
        st.metric("处理失败的文件夹", len(folder_status['error']))
    
    with st.expander("查看详细处理状态", expanded=True):
        if folder_status['success']:
            st.subheader("✅ 成功处理的文件夹")
            for folder_name, processed, failed in folder_status['success']:
                st.markdown(f"""
                - **{folder_name}**
                    - 成功处理: {processed} 个文件
                    - 跳过: {failed} 个文件
                """)
        
        if folder_status['empty']:
            st.subheader("⚠️ 空文件夹")
            for folder_name in folder_status['empty']:
                st.markdown(f"- **{folder_name}**: 未找到Word文件")
        
        if folder_status['error']:
            st.subheader("❌ 处理失败的文件夹")
            for folder_name, error_msg in folder_status['error']:
                st.markdown(f"- **{folder_name}**: {error_msg}")
        
        if folder_status['file_errors']:
            st.subheader("📄 文件处理错误详情")
            for folder_name, failed_files in folder_status['file_errors'].items():
                with st.expander(f"{folder_name} - {len(failed_files)}个文件失败"):
                    for file_name in failed_files:
                        st.markdown(f"- {file_name}")

    st.markdown("---")

    if all_tables:
        merged_all = pd.concat(all_tables, ignore_index=True)
        excel_buffer = BytesIO()
        merged_all.to_excel(excel_buffer, index=False)
        excel_buffer.seek(0)

    if all_warning_samples:
        combined_warnings = pd.concat(all_warning_samples, ignore_index=True)
        status_summary = combined_warnings['状态'].value_counts()
        warning_buffer = BytesIO()
        combined_warnings.to_excel(warning_buffer, index=False)
        warning_buffer.seek(0)

    tab1, tab2 = st.tabs(["数据汇总", "预警信息"])
    
    with tab1:
        if merged_all is not None:
            st.subheader("所有表格汇总")
            st.dataframe(merged_all, use_container_width=True)
            
            col1, _ = st.columns([1, 3])
            with col1:
                st.download_button(
                    label="📥 下载汇总表格",
                    data=excel_buffer,
                    file_name=f'汇总表格_{current_time}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                )
        else:
            st.info("暂无数据")

    with tab2:
        if combined_warnings is not None:
            st.subheader("预警信息")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("已超期样品", f"{status_summary.get('❌ 已超期', 0)}个")
            with col2:
                st.metric("接近超期样品", f"{status_summary.get('⚠️ 接近超期', 0)}个")
            with col3:
                st.metric("总预警样品", f"{len(combined_warnings)}个")
            
            st.markdown("---")
            st.dataframe(combined_warnings, use_container_width=True)
            
            col1, _ = st.columns([1, 3])
            with col1:
                st.download_button(
                    label="📥 下载预警信息",
                    data=warning_buffer,
                    file_name=f'预警信息_{current_time}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                )
        else:
            st.info("暂无预警信息")

    if st.button("清除所有缓存数据"):
        st.session_state.warning_data = {}
        st.session_state.merged_tables = {}
        st.rerun()

def main():
    st.set_page_config(
        page_title="样品预警系统",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    st.markdown("""
        <style>
        .main {
            padding: 2rem;
        }
        .stButton > button {
            width: 100%;
            border-radius: 4px;
            padding: 0.5rem;
            background-color: #f0f2f6;
            border: 1px solid #e0e3e9;
        }
        .stButton > button:hover {
            background-color: #e0e3e9;
        }
        .st-emotion-cache-16idsys {
            font-size: 1rem;
            padding: 0.5rem;
        }
        .st-emotion-cache-1y4p8pa {
            max-width: 100%;
        }
        .st-emotion-cache-1wivap2 {
            padding: 1rem;
            border-radius: 4px;
            background-color: white;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        }
        </style>
    """, unsafe_allow_html=True)

    st.sidebar.title("功能选择")
    page = st.sidebar.radio("", ["表格处理", "配置管理"], label_visibility="collapsed")
    
    if page == "表格处理":
        process_tables()
    else:
        config_editor()

if __name__ == "__main__":
    main() 