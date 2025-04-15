import streamlit as st
import pandas as pd
import numpy as np
import io
import os
from datetime import datetime

# 设置页面标题
st.set_page_config(page_title="电磁炉物料清单管理系统", layout="wide")

# 页面标题
st.title("电磁炉物料清单管理系统")

# 加载Excel文件
@st.cache_data(ttl=60)  # 缓存1分钟
def load_excel_file(file):
    try:
        df = pd.read_excel(file)
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
        ["上传Excel文件", "使用示例数据"]
    )
    
    # 添加重置按钮
    if st.button("重置应用"):
        reset_app()
    
    # 帮助信息
    st.markdown("---")
    st.subheader("帮助信息")
    st.markdown("""
    **使用步骤:**
    1. 上传父件和子件Excel文件
    2. 选择要生产的电磁炉型号
    3. 输入生产数量
    4. 生成物料需求计划
    5. 下载Excel结果
    
    **文件格式要求:**
    - 父件文件需包含: 物料清单编码, 父件商品
    - 子件文件需包含: 物料清单编码, 子件商品, 规格型号, 需用数量, 成本单价, 成本金额, 默认供应商
    """)

# 主页面逻辑
if upload_option == "上传Excel文件":
    st.subheader("上传Excel文件")
    st.info("请上传包含物料清单的Excel文件。文件大小限制: 200MB")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 上传物料清单父件文件")
        parent_file = st.file_uploader("选择父件Excel文件", type=['xlsx', 'xls'], key='parent', 
                                      accept_multiple_files=False, help="请上传包含物料清单父件的Excel文件")
        
    with col2:
        st.markdown("### 上传物料清单子件文件")
        child_file = st.file_uploader("选择子件Excel文件", type=['xlsx', 'xls'], key='child', 
                                     accept_multiple_files=False, help="请上传包含物料清单子件的Excel文件")
    
    # 示例文件格式
    with st.expander("查看文件格式要求"):
        st.markdown("""
        #### 父件文件必须包含以下列:
        - **物料清单编码**: 唯一标识符，用于关联父子件
        - **父件商品**: 电磁炉型号名称
        
        #### 子件文件必须包含以下列:
        - **物料清单编码**: 与父件的物料清单编码对应
        - **子件商品**: 零部件名称
        - **规格型号**: 零部件规格
        - **需用数量**: 每台电磁炉需要的零部件数量
        - **成本单价**: 零部件单价
        - **成本金额**: 单个零部件的总成本
        - **默认供应商**: 供应商名称
        """)
        
        # 添加示例图片或更详细的格式说明
        st.markdown("##### 文件格式示例:")
        st.code("""
父件文件示例:
物料清单编码 | 父件商品           | 生产数量 | 成本金额
-----------|------------------|---------|--------
0000072    | 5KW380V双平旋钮   | 1       | 712.77
0000075    | 5KW380V双平磁控   | 1       | 865.32

子件文件示例:
物料清单编码 | 子件商品 | 规格型号      | 需用数量 | 成本单价 | 成本金额 | 默认供应商
-----------|---------|--------------|---------|--------|---------|----------
0000072    | SZ-3868 | 500*275*20mm | 3       | 3.00   | 9.00    | 供应商A
0000072    | 护边    | 500长（家用）  | 6       | 1.00   | 6.00    | 供应商B
        """)
    
    # 如果通过文件上传器上传了文件
    if parent_file is not None and child_file is not None:
        with st.spinner("正在处理上传的文件..."):
            # 读取父件数据
            df_parent, parent_error = load_excel_file(parent_file)
            # 读取子件数据
            df_child, child_error = load_excel_file(child_file)
            
            if df_parent is not None and df_child is not None:
                # 展示原始数据
                st.subheader("父件数据预览")
                st.dataframe(df_parent.head(), use_container_width=True)
                
                st.subheader("子件数据预览")
                st.dataframe(df_child.head(), use_container_width=True)
                
                # 验证数据格式
                required_parent_columns = ["物料清单编码", "父件商品"]
                required_child_columns = ["物料清单编码", "子件商品", "需用数量", "成本单价", "成本金额"]
                
                missing_parent_cols = [col for col in required_parent_columns if col not in df_parent.columns]
                missing_child_cols = [col for col in required_child_columns if col not in df_child.columns]
                
                if missing_parent_cols or missing_child_cols:
                    if missing_parent_cols:
                        st.error(f"父件文件缺少必要列: {', '.join(missing_parent_cols)}")
                    if missing_child_cols:
                        st.error(f"子件文件缺少必要列: {', '.join(missing_child_cols)}")
                else:
                    # 存储处理后的数据
                    st.session_state.processed_data = {
                        "parent": df_parent,
                        "child": df_child
                    }
                    
                    st.success("文件上传成功！请继续进行生产计划设置。")
            else:
                if parent_error:
                    st.error(f"无法加载父件文件: {parent_error}")
                if child_error:
                    st.error(f"无法加载子件文件: {child_error}")

else:
    # 使用示例数据
    try:
        demo_parent = """
物料清单编码,父件商品,生产数量,成本金额
0000072,5KW380V双平旋钮(5000W),1,712.77
0000075,5KW380V双平磁控(5000W),1,865.32
0000136,出口双电磁 110V,1,188.67
0000151,米洲5000W220V凹面旋钮（5000W）,1,195.52
0000150,米洲5000W220V平面旋钮（5000W）,1,187.38
        """
        
        demo_child = """
物料清单编码,子件商品,规格型号,需用数量,成本单价,成本金额,默认供应商
0000072,SZ-3868,,3,3.00,9.00,供应商A
0000072,护边500长（家用双灶）,,6,1.00,6.00,供应商B
0000072,泡沫块配3868,500*275*20mm,3,0.55,1.65,供应商C
0000075,SZ-3868,,3,3.00,9.00,供应商A
0000075,护边500长（家用双灶）,,6,1.00,6.00,供应商B
0000075,磁控开关,,2,15.00,30.00,供应商D
0000136,控制器-110V,,1,16.30,16.30,供应商E
0000136,电源线,,1,3.50,3.50,供应商F
0000150,控制器-220V,,1,18.50,18.50,供应商E
0000150,电源线,,1,3.80,3.80,供应商F
0000151,控制器-220V高配,,1,21.50,21.50,供应商E
0000151,电源线,,1,3.80,3.80,供应商F
        """
        
        df_parent = pd.read_csv(io.StringIO(demo_parent.strip()))
        df_child = pd.read_csv(io.StringIO(demo_child.strip()))
        
        # 展示原始数据
        st.subheader("父件数据预览（示例数据）")
        st.dataframe(df_parent, use_container_width=True)
        
        st.subheader("子件数据预览（示例数据）")
        st.dataframe(df_child, use_container_width=True)
        
        # 存储处理后的数据
        st.session_state.processed_data = {
            "parent": df_parent,
            "child": df_child
        }
        
        st.success("示例数据加载成功！请继续进行生产计划设置。")
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
        with st.spinner("正在生成物料需求计划..."):
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

# 页脚
st.markdown("---")
with st.container():
    cols = st.columns([1, 2, 1])
    with cols[1]:
        st.caption("电磁炉物料清单管理系统 © 2023")
        st.caption("有问题请联系管理员") 
