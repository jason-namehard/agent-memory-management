# -*- coding: utf-8 -*-
"""归档提醒脚本：周三/周日定时提醒用户完成 session 归档"""
import datetime

def main():
    wd = datetime.datetime.now().weekday()  # 0=Mon 2=Wed 6=Sun
    if wd == 6:  # 周日
        print("📋 周日归档提醒：请先完成归档（总结+记忆迁移+开新 session）再开始其它任务。5 天周期已到，本周 session 任务清单待裁决。")
    elif wd == 2:  # 周三
        print("📋 周三归档提醒（5 天周期）：请对当前 session 任务清单做裁决（继承/删除/压缩），完成记忆迁移，必要时开新 session。")
    else:
        print("")  # 静默

if __name__ == "__main__":
    main()
