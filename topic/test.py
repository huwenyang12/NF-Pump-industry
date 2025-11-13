# 原始收货人信息
str = "送货至/江苏省/无锡市/惠山区前洲街道谢印路无锡创新低温环模设备科技有限公司/钱嘉诚/15852796275"
text = str.strip()
total_len = len(text)
# 按每行最多 21 个字符自动换行
if total_len > 21:
    new_text = ""
    for i in range(0, total_len, 21):
        new_text += text[i:i+21]
        if i + 21 < total_len:
            new_text += "\n"
    text = new_text

print(text)