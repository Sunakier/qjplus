# qjplus v4-202510

Python 实现的基于 DrissionPage 浏览器自动化 + 协议实现的 青骄第二课堂自动化解决方案  
本项目是一个 v3 的重构版本

## 如何使用

1. **安装依赖**:
   ```bash
   pip install -r requirements.txt
   ```

2. **配置**:
   - 根据需要修改 `config/config.json` 文件中的配置，其中 `browser_path` 、 `user_data_files` 和 `willFinishListIfLess` 应被格外注意

3. **准备用户数据**:
   - 将包含用户账号信息（学生姓名、账号、密码、年级）的表格文件（.xls, .xlsx, .csv）放入 `user_data_files` 指定的目录中。

4. **运行**:
   ```bash
   python src/main.py
   ```

## 项目结构

```
.
├── config/               # 配置文件目录
├── logs/                 # 日志文件目录
├── src/                  # 源代码目录
│   ├── browser_manager.py
│   ├── config_manager.py
│   ├── data_processor.py
│   ├── login_manager.py
│   ├── log_manager.py
│   ├── main.py
│   ├── report_generator.py
│   ├── task_manager.py
│   └── utils.py
├── tests/                # 测试代码目录
├── users/                # 用户数据目录
├── README.md
└── requirements.txt

## 联络我

Bilbili 演示视频(v3):

[BV1mP411N7Qe](https://www.bilibili.com/video/BV1mP411N7Qe)

[BV1Aw411N7bb](https://www.bilibili.com/video/BV1Aw411N7bb)

Mail: [lazyerpaper@qq.com](mailto:lazyerpaper@qq.com)

WeChat: lazyerpaper

Telegram: [Wuqibor](https://t.me/Wuqibor)