"""
generate_sample_data.py
────────────────────────
生成包含 5 条模拟记录的测试 Excel 文件。
运行方式：python generate_sample_data.py
"""

from pathlib import Path
import pandas as pd

sample_records = [
    {
        "Case Contact Name": "姓名：陈建国\n日期：2025.01.05-2025.01.06\n渠道：美团",
        "Location": "521",
        "CaseNumber": "100001",
        "Member": "NIL",
        "Description": "客人反映空调噪音影响睡眠\nGuest complained about air conditioning noise affecting sleep",
        "Resolution Notes": (
            "22:35客人致电前台，反映521房间空调运行时有规律性的嗡嗡声，严重影响睡眠，"
            "要求酒店解决。前台联系工程部，工程师23:00到房检查，确认为空调过滤网积灰导致振动。"
            "当场清洁过滤网后噪音消失，客人表示满意，工程师离开后客人再次确认无异常。"
            "次日早餐已安排赠送以示歉意，客人表示感谢。"
        ),
    },
    {
        "Case Contact Name": "姓名：Sarah Mitchell（英国籍）\n日期：2025.01.05-2025.01.08\n渠道：Booking.com",
        "Location": "1203",
        "CaseNumber": "100002",
        "Member": "Gold",
        "Description": "外籍客人对早餐服务不满意\nForeign guest dissatisfied with breakfast service",
        "Resolution Notes": (
            "09:15，Sarah女士来到前台，情绪较为激动，反映在餐厅用早餐时服务员态度冷漠，"
            "点餐两次被忽视，等待超过20分钟仍无人响应，最终只能自取食物。"
            "Sarah表示这与五星级酒店的服务承诺严重不符，要求酒店作出解释。"
            "餐饮部经理James随即向客人致歉，询问当值服务员，确认因早餐高峰期人手不足，"
            "导致外籍区域服务疏漏。当日免除早餐费用，并为其余入住期间升级午餐权益，Sarah接受道歉。"
        ),
    },
    {
        "Case Contact Name": "姓名：王磊\n日期：2025.01.06-2025.01.07\n渠道：携程",
        "Location": "309/310",
        "CaseNumber": "100003",
        "Member": "Platinum",
        "Description": "联通房隔音问题导致客人投诉\nConnecting room noise complaint",
        "Resolution Notes": (
            "深夜23:40，王先生致电前台，愤怒反映其309联通房隔壁310房间持续发出喧嚣声，"
            "包括聚会音乐和大声谈话，严重影响其休息。前台立即联系保安巡查，"
            "确认310房间共7人，属超员入住且存在扰邻行为。"
            "保安礼貌但坚定地要求310房客人降低音量并送走多余人员，"
            "期间310客人情绪激动与保安发生口角，经安保主管出面调解，24:10恢复安静。"
            "王先生对处理速度表示认可，次日前台主动致电确认睡眠质量，客人表示满意，"
            "酒店赠送积分补偿。"
        ),
    },
    {
        "Case Contact Name": "姓名：刘晓燕\n日期：2025.01.07-2025.01.09\n渠道：官网直订",
        "Location": "Lobby",
        "CaseNumber": "100004",
        "Member": "Ambassador",
        "Description": "贵重物品遗落大堂\nGuest left valuables in lobby",
        "Resolution Notes": (
            "15:30，前台接到客人刘晓燕电话，告知其离店约1小时后发现手提包遗落在大堂沙发区，"
            "包内有钱包（现金约3000元及信用卡）、护照及部分首饰。"
            "安保调取14:00-14:30大堂监控，发现客人离开时包遗落在沙发，"
            "约5分钟后由大堂保洁员发现并第一时间上交前台，现完整保管于前台保险箱。"
            "酒店立即联系客人，客人打车返回后核对物品完好无缺，情绪从紧张转为感激，"
            "对酒店诚信保管行为高度称赞，并表示会在平台给予好评。"
        ),
    },
    {
        "Case Contact Name": "姓名：赵宏伟\n日期：2025.01.08\n渠道：团购",
        "Location": "餐厅",
        "CaseNumber": "100005",
        "Member": "NIL",
        "Description": "餐厅点餐异物事件\nForeign object found in food",
        "Resolution Notes": (
            "19:45，赵先生在酒店中餐厅用餐时，在例汤中发现一根约3cm的黑色异物（疑似头发）。"
            "赵先生立即招呼服务员，情绪激动，当着周围客人大声质问食品卫生问题，"
            "引发其他用餐客人关注。餐厅领班孙小姐迅速到场，第一时间诚挚致歉，"
            "将餐品撤换并请客人移至包厢以避开公众视线，赠送当日消费全免并提供特色甜品。"
            "行政总厨随后亲自到场致歉并说明将彻查后厨操作规范，赵先生情绪逐渐平息，"
            "离店前表示理解，但希望酒店切实改进。事后厨房已对所有员工开展操作规范专项检查。"
        ),
    },
]

df = pd.DataFrame(sample_records)

output_path = Path(__file__).parent / "sample_data" / "sample_logbook.xlsx"
output_path.parent.mkdir(exist_ok=True)

with pd.ExcelWriter(str(output_path), engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="GSM Log", index=False)

print(f"测试文件已生成：{output_path}")
print(f"共 {len(df)} 条记录")
