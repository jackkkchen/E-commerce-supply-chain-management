import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import glob
from datetime import datetime

# 设置页面标题
st.set_page_config(page_title="电磁炉物料清单管理系统", layout="wide")

# 页面标题
st.title("电磁炉物料清单管理系统")

# 扫描工作目录中的Excel文件
@st.cache_data(ttl=60)  # 缓存1分钟
def scan_excel_files():
    excel_files = glob.glob("*.xlsx") + glob.glob("*.xls")
    return excel_files

# 加载Excel文件
@st.cache_data(ttl=60)  # 缓存1分钟
def load_excel_file(file_path):
    try:
        df = pd.read_excel(file_path)
        return df, None
    except Exception as e:
        return None, str(e)

# 初始化session_state
if "processed_data" not in st.session_state:
    st.session_state.processed_data = None
if "production_plan" not in st.session_state:
    st.session_state.production_plan = None

# 重置功能函数
def reset_app():
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.experimental_rerun()

# 侧边栏
with st.sidebar:
    st.header("操作面板")
    upload_option = st.radio(
        "选择上传方式",
        ["上传Excel文件", "使用示例数据", "从工作目录加载"]
    )
    
    # 添加重置按钮
    if st.button("重置应用"):
        reset_app()

# 主页面逻辑
if upload_option == "上传Excel文件":
    st.write("如果文件上传按钮无响应，请尝试清除浏览器缓存或使用其他浏览器，或者选择'从工作目录加载'选项。")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("上传物料清单父件文件")
        parent_file = st.file_uploader("选择父件Excel文件", type=['xlsx', 'xls'], key='parent', 
                                      accept_multiple_files=False, help="请上传包含物料清单父件的Excel文件")
        
        # 添加替代文件路径输入方式
        parent_file_path = st.text_input("或输入父件文件路径:", placeholder="例如: ./物料清单父件.xlsx")
        if parent_file_path and os.path.exists(parent_file_path) and parent_file is None:
            df_parent, error = load_excel_file(parent_file_path)
            if df_parent is not None:
                st.success(f"从路径 {parent_file_path} 成功加载父件数据")
                # 存储临时数据到session state
                st.session_state.df_parent = df_parent
            else:
                st.error(f"无法从路径加载文件: {error}")
        
    with col2:
        st.subheader("上传物料清单子件文件")
        child_file = st.file_uploader("选择子件Excel文件", type=['xlsx', 'xls'], key='child', 
                                     accept_multiple_files=False, help="请上传包含物料清单子件的Excel文件")
        
        # 添加替代文件路径输入方式
        child_file_path = st.text_input("或输入子件文件路径:", placeholder="例如: ./物料清单父子件.xlsx")
        if child_file_path and os.path.exists(child_file_path) and child_file is None:
            df_child, error = load_excel_file(child_file_path)
            if df_child is not None:
                st.success(f"从路径 {child_file_path} 成功加载子件数据")
                # 存储临时数据到session state
                st.session_state.df_child = df_child
            else:
                st.error(f"无法从路径加载文件: {error}")
    
    # 如果通过文件上传器上传了文件
    if parent_file is not None and child_file is not None:
        try:
            # 读取父件数据
            df_parent = pd.read_excel(parent_file)
            # 读取子件数据
            df_child = pd.read_excel(child_file)
            
            # 展示原始数据
            st.subheader("父件数据预览")
            st.dataframe(df_parent.head())
            
            st.subheader("子件数据预览")
            st.dataframe(df_child.head())
            
            # 存储处理后的数据
            st.session_state.processed_data = {
                "parent": df_parent,
                "child": df_child
            }
            
            st.success("文件上传成功！请继续进行生产计划设置。")
        except Exception as e:
            st.error(f"处理文件时出错: {e}")
    
    # 如果通过文件路径输入了文件
    elif hasattr(st.session_state, 'df_parent') and hasattr(st.session_state, 'df_child'):
        # 展示原始数据
        st.subheader("父件数据预览")
        st.dataframe(st.session_state.df_parent.head())
        
        st.subheader("子件数据预览")
        st.dataframe(st.session_state.df_child.head())
        
        # 存储处理后的数据
        st.session_state.processed_data = {
            "parent": st.session_state.df_parent,
            "child": st.session_state.df_child
        }
        
        if st.button("确认使用这些数据"):
            st.success("数据加载成功！请继续进行生产计划设置。")

elif upload_option == "从工作目录加载":
    st.subheader("从工作目录加载Excel文件")
    
    # 扫描工作目录中的Excel文件
    excel_files = scan_excel_files()
    
    if not excel_files:
        st.warning("当前工作目录中未找到Excel文件。")
    else:
        st.write(f"找到 {len(excel_files)} 个Excel文件:")
        
        # 显示可用文件列表
        for i, file in enumerate(excel_files):
            st.write(f"{i+1}. {file}")
        
        # 选择父件和子件文件
        parent_file_idx = st.selectbox("选择父件文件", 
                                      options=list(range(len(excel_files))), 
                                      format_func=lambda x: excel_files[x])
        
        child_file_idx = st.selectbox("选择子件文件", 
                                     options=list(range(len(excel_files))), 
                                     format_func=lambda x: excel_files[x])
        
        if st.button("加载选中的文件"):
            parent_path = excel_files[parent_file_idx]
            child_path = excel_files[child_file_idx]
            
            df_parent, parent_error = load_excel_file(parent_path)
            df_child, child_error = load_excel_file(child_path)
            
            if df_parent is not None and df_child is not None:
                # 展示原始数据
                st.subheader("父件数据预览")
                st.dataframe(df_parent.head())
                
                st.subheader("子件数据预览")
                st.dataframe(df_child.head())
                
                # 存储处理后的数据
                st.session_state.processed_data = {
                    "parent": df_parent,
                    "child": df_child
                }
                
                st.success("文件加载成功！请继续进行生产计划设置。")
            else:
                if parent_error:
                    st.error(f"无法加载父件文件: {parent_error}")
                if child_error:
                    st.error(f"无法加载子件文件: {child_error}")

else:
    # 使用示例数据
    try:
        if os.path.exists("物料清单父件.xlsx") and os.path.exists("物料清单父子件.xlsx"):
            # 读取示例父件数据
            df_parent, _ = load_excel_file("物料清单父件.xlsx")
            # 读取示例子件数据
            df_child, _ = load_excel_file("物料清单父子件.xlsx")
            
            if df_parent is not None and df_child is not None:
                # 展示原始数据
                st.subheader("父件数据预览（示例数据）")
                st.dataframe(df_parent.head())
                
                st.subheader("子件数据预览（示例数据）")
                st.dataframe(df_child.head())
                
                # 存储处理后的数据
                st.session_state.processed_data = {
                    "parent": df_parent,
                    "child": df_child
                }
                
                st.success("示例数据加载成功！请继续进行生产计划设置。")
            else:
                st.warning("示例数据文件加载失败。")
        else:
            st.warning("示例数据文件不存在，请选择上传文件方式。")
    except Exception as e:
        st.error(f"加载示例数据时出错: {e}")

# 生产计划设置
if st.session_state.processed_data is not None:
    st.header("设置生产计划")
    
    parent_data = st.session_state.processed_data["parent"]
    
    # 获取所有父件商品名称
    parent_products = parent_data["父件商品"].unique().tolist()
    
    # 创建生产计划设置界面
    selected_product = st.selectbox("选择要生产的电磁炉型号", parent_products)
    production_quantity = st.number_input("生产数量", min_value=1, value=10, step=1)
    
    if st.button("生成物料需求计划"):
        try:
            # 筛选选定的父件
            selected_parent = parent_data[parent_data["父件商品"] == selected_product].iloc[0]
            parent_code = selected_parent["物料清单编码"]
            
            # 筛选对应的子件
            child_data = st.session_state.processed_data["child"]
            
            # 查找与选定父件相关的所有子件
            selected_children = child_data[child_data["物料清单编码"] == parent_code].copy()
            
            if selected_children.empty:
                st.error(f"未找到与'{selected_product}'相关的子件数据。")
            else:
                # 强制转换需用数量列为数值型
                try:
                    selected_children["需用数量"] = pd.to_numeric(selected_children["需用数量"], errors='coerce')
                    selected_children["成本单价"] = pd.to_numeric(selected_children["成本单价"], errors='coerce')
                    selected_children["成本金额"] = pd.to_numeric(selected_children["成本金额"], errors='coerce')
                except Exception as e:
                    st.warning(f"数据类型转换警告: {e}，某些计算可能不准确")
                
                # 计算总需求量
                selected_children["需用数量_总计"] = selected_children["需用数量"] * production_quantity
                selected_children["成本金额_总计"] = selected_children["成本金额"] * production_quantity
                
                # 准备输出数据
                try:
                    output_columns = ["子件商品", "规格型号", "需用数量", "成本单价", "成本金额", "需用数量_总计", "成本金额_总计", "默认供应商"]
                    available_columns = [col for col in output_columns if col in selected_children.columns]
                    output_data = selected_children[available_columns].copy()
                    
                    # 如果缺少某些列，添加空列
                    for col in output_columns:
                        if col not in output_data.columns:
                            output_data[col] = np.nan
                            
                except Exception as e:
                    st.warning(f"选择列时出现问题: {e}，将使用所有可用列")
                    output_data = selected_children.copy()
                
                # 添加总计行
                total_cost = output_data["成本金额_总计"].sum()
                total_row_data = {}
                
                # 初始化所有列
                for col in output_data.columns:
                    total_row_data[col] = [""] if col != "成本金额_总计" else [total_cost]
                
                # 设置特殊列
                total_row_data["子件商品"] = [""]
                total_row_data["规格型号"] = [""]
                if "需用数量" in total_row_data:
                    total_row_data["需用数量"] = [np.nan]
                if "成本单价" in total_row_data:
                    total_row_data["成本单价"] = [np.nan]
                if "成本金额" in total_row_data:
                    total_row_data["成本金额"] = [np.nan]
                if "需用数量_总计" in total_row_data:
                    total_row_data["需用数量_总计"] = ["成本金额汇总："]
                if "默认供应商" in total_row_data:
                    total_row_data["默认供应商"] = [""]
                
                total_row = pd.DataFrame(total_row_data)
                
                # 转换所有列为对象类型，避免Arrow转换问题
                for col in total_row.columns:
                    total_row[col] = total_row[col].astype('object')
                
                # 将所有输出数据列转换为object类型避免Arrow问题
                for col in output_data.columns:
                    output_data[col] = output_data[col].astype('object')
                
                # 合并数据
                try:
                    output_data = pd.concat([output_data, total_row], ignore_index=True)
                except Exception as e:
                    st.error(f"合并数据时出错: {e}")
                    st.write("总计行:", total_row)
                    st.write("输出数据:", output_data.dtypes)
                
                # 存储生产计划数据
                st.session_state.production_plan = {
                    "product": selected_product,
                    "quantity": production_quantity,
                    "output_data": output_data
                }
                
                # 显示生成的物料需求计划
                st.subheader(f"{selected_product} - 生产数量: {production_quantity}台")
                
                # 使用st.table代替st.dataframe来避免Arrow兼容性问题
                st.table(output_data)
                
                # 提供下载链接
                st.success("物料需求计划生成成功！")
                
                # 按供应商分类
                if "默认供应商" in selected_children.columns:
                    if st.checkbox("按供应商分类显示"):
                        st.subheader("按供应商分类的物料需求")
                        suppliers = selected_children["默认供应商"].dropna().unique()
                        
                        for supplier in suppliers:
                            supplier_data = selected_children[selected_children["默认供应商"] == supplier].copy()
                            supplier_total = supplier_data["成本金额_总计"].sum()
                            
                            st.write(f"供应商: {supplier} - 总成本: {supplier_total:.2f}")
                            display_cols = ["子件商品", "规格型号", "需用数量_总计", "成本单价", "成本金额_总计"]
                            available_display_cols = [col for col in display_cols if col in supplier_data.columns]
                            st.table(supplier_data[available_display_cols])
        
        except Exception as e:
            st.error(f"生成物料需求计划时出错: {e}")
            import traceback
            st.error(traceback.format_exc())

# 导出数据
if st.session_state.production_plan is not None:
    st.header("导出数据")
    
    output_data = st.session_state.production_plan["output_data"]
    product_name = st.session_state.production_plan["product"]
    quantity = st.session_state.production_plan["quantity"]
    
    # 创建Excel文件用于下载
    output = io.BytesIO()
    
    try:
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            output_data.to_excel(writer, index=False, sheet_name="物料需求计划")
        
        # 提供下载链接
        current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{product_name}_物料需求计划_{quantity}台_{current_time}.xlsx"
        
        st.download_button(
            label="下载Excel文件",
            data=output.getvalue(),
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"生成Excel文件时出错: {e}")
        st.write("尝试保存到本地文件...")
        
        # 尝试直接保存到本地文件
        try:
            local_filename = f"{product_name}_物料需求计划_{quantity}台_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            output_data.to_excel(local_filename, index=False, sheet_name="物料需求计划")
            st.success(f"文件已保存到本地: {os.path.abspath(local_filename)}")
        except Exception as e2:
            st.error(f"保存到本地文件时出错: {e2}")

# 页脚
st.markdown("---")
st.caption("电磁炉物料清单管理系统 © 2023") 
