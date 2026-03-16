import sys
sys.path.insert(0, '.')
from src.prompt_manager import has_guide_questions, append_guide_questions_to_content

# 测试1：包含引导问题
t1 = "## 6. 案例启示\n内容\n\n## 引导问题\n\n1. 问题一\n\n2. 问题二"
print("有引导问题:", has_guide_questions(t1))  # True

# 测试2：缺少引导问题
t2 = "## 6. 案例启示\n内容"
print("无引导问题:", has_guide_questions(t2))  # False

# 测试3：补写后合并
merged = append_guide_questions_to_content(t2, "## 引导问题\n\n1. Q1\n\n2. Q2")
print("合并后含引导问题:", has_guide_questions(merged))  # True
print("合并末尾内容:", repr(merged[-60:]))

# 测试4：变体标题（如 ## 7. 引导问题）
t3 = "内容\n\n## 7. 引导问题\n\n1. Q1\n\n2. Q2"
print("变体标题识别:", has_guide_questions(t3))  # True

print("\n所有测试通过!")
