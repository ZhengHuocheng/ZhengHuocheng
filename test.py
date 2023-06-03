import streamlit as st
st.title("问答系统")
st.write("输入一个问题和一段文本，获取问题的答案")
# 用户输入问题和文本
question = st.text_input("请输入你的问题")
text = st.text_area("请输入文本")
# 提交按钮
if st.button("提交"):
    # 对输入的问题和文本进行问答
    answer = "你好"

    # 显示答案
    st.write("问题：", "你好帅")
    st.write("答案：", "是的")
    st.write("置信度：", "100")
