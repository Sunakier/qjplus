# DrissionPage完整文档

## 目录

- [SessionPage](#SessionPage)
- [下载](#下载)
- [入门](#入门)
- [其他](#其他)
- [教程](#教程)
- [浏览器控制](#浏览器控制)
- [版本](#版本)
- [特性](#特性)

## <a name="其他"></a>其他

* 🧰 进阶使用
* ⚙️ 数据读取加速

本页总览

⚙️ 数据读取加速
=========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

本节演示一个能够大幅加快浏览器页面数据采集的黑科技。

✅️️ 示例[​](#️️-示例 "✅️️ 示例的直接链接")
-------------------------------

我们找一个比较大的页面来演示，比如网页首页：<https://www.163.com>

我们数一下这个网页内的`<a>`元素数量：

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.get('https://www.163.com')  
print(len(tab('t:body').eles('t:a')))
```

**输出：**

```
1613
```

嗯，数量不少，可以看出效果。

假如现在我们的任务是打印所有链接的文本，常规做法是遍历所有元素，然后打印。

这里引入本库作者写的一个计时工具，可以标记一段代码运行时间。您也可以用其它方法计时。

```
from DrissionPage import Chromium  
from TimePinner import Pinner  # 导入计时工具  
  
pinner = Pinner()  # 创建计时器对象  
tab = Chromium().latest_tab  
tab.get('https://www.163.com')  
  
pinner.pin()  # 标记开始记录  
  
# 获取所有链接对象并遍历  
links = tab('t:body').eles('t:a')  
for lnk in links:  
    print(lnk.text)  
  
pinner.pin('用时')  # 记录并打印时间
```

**输出：**

```
0.0  
  
网络大过年_网易政务_网易网  
网易首页  
...中间省略...  
不良信息举报 Complaint Center  
廉正举报  
用时：4.057772700001806
```

用时 4 秒。

现在，我们稍微修改一个小小的地方。

把`page('t:body').eles('t:a')`改成`page('t:body').s_eles('t:a')`，然后再执行一次。

```
from DrissionPage import Chromium  
from TimePinner import Pinner  # 导入计时工具  
  
pinner = Pinner()  # 创建计时器对象  
tab = Chromium().latest_tab  
tab.get('https://www.163.com')  
  
pinner.pin()  # 标记开始记录  
  
# 获取所有链接对象并遍历  
links = tab('t:body').s_eles('t:a')  
for lnk in links:  
    print(lnk.text)  
  
pinner.pin('用时')  # 记录并打印时间
```

**输出：**

```
0.0  
  
网络大过年_网易政务_网易网  
网易首页  
...中间省略...  
不良信息举报 Complaint Center  
廉正举报  
用时：0.2797656000002462
```

神奇不？原来 4 秒的采集时间现在只需 0.28 秒。

---

✅️️ 解读[​](#️️-解读 "✅️️ 解读的直接链接")
-------------------------------

`s_eles()`与`eles()`的区别在于前者会把整个页面或动态元素转变成一个静态元素，再在其中获取下级元素或信息。

因为静态元素是纯文本的，没有各种属性、交互等消耗资源的部分，所以运行速度非常快。

作者曾经采集过一个非常复杂的页面，动态元素用时 30 秒，转静态元素就只要 0.X 秒，加速效果非常明显。

我们可以获取页面中内容容器（示例中的`<body>`），把它转换成静态元素，再在其中获取信息。

当然，静态元素没有交互功能，它只是副本，也不会影响原来的动态元素。

说明

一个页面中不用反复使用`s_ele()`，通常只要使用一次，获取最高级的容器元素或者页面对象本身的静态副本，然后在这个副本中查找元素。
反复使用的话会因为资源消耗较大导致不稳定和浪费时间。

[上一页

⚙️ 异常的使用](/advance/errors)[下一页

⚙️ 打包程序](/advance/packaging)

* [✅️️ 示例](#️️-示例)
* [✅️️ 解读](#️️-解读)

* 🧰 进阶使用
* ⚙️ 命令行的使用

本页总览

⚙️ 命令行的使用
=========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

DrissionPage 提供一些便捷的命令行命令，用于基本设置，以取代有时需要的临时配置文件。

命令行主命令为`dp`，形式为：

```
dp 命令全称或缩写 <参数>
```

✅️️ 设置浏览器路径[​](#️️-设置浏览器路径 "✅️️ 设置浏览器路径的直接链接")
----------------------------------------------

| 全称 | 缩写 | 参数 | 说明 |
| --- | --- | --- | --- |
| --set-browser-path | -p | 浏览器路径 | 设置配置文件中的浏览器路径 |

**示例：**

```
# 完整写法  
dp --set-browser-path "D:\chrome\Chrome.exe"  
  
# 简略写法  
dp -p "D:\chrome\Chrome.exe"
```

✅️️ 设置用户数据路径[​](#️️-设置用户数据路径 "✅️️ 设置用户数据路径的直接链接")
-------------------------------------------------

| 全称 | 缩写 | 参数 | 说明 |
| --- | --- | --- | --- |
| --set-user-path | -u | 用户数据文件夹路径 | 设置配置文件中的用户数据路径 |

**示例：**

```
# 完整写法  
dp --set-user-path D:\chrome\user_data  
  
# 简略写法  
dp -u D:\chrome\user_data
```

  

✅️️ 复制默认 ini 文件到当前路径[​](#️️-复制默认-ini-文件到当前路径 "✅️️ 复制默认 ini 文件到当前路径的直接链接")
-------------------------------------------------------------------------

| 全称 | 缩写 | 参数 | 说明 |
| --- | --- | --- | --- |
| --configs-to-here | -c | 无 | 复制默认配置文件到当前路径 |

**示例：**

```
# 完整写法  
dp --configs-to-here  
  
# 简略写法  
dp -c
```

✅️️ 启动浏览器[​](#️️-启动浏览器 "✅️️ 启动浏览器的直接链接")
----------------------------------------

此命令用于启动浏览器，等待程序接管。

| 全称 | 缩写 | 参数 | 说明 |
| --- | --- | --- | --- |
| --launch-browser | -l | 端口号 | 启动浏览器，传入端口号，0表示用配置文件中的值 |

**示例：**

```
# 完整写法  
dp --launch-browser 9333  
  
# 简略写法  
dp -l 0
```

[上一页

⚙️ 全局设置](/advance/settings)[下一页

⚙️ 异常的使用](/advance/errors)

* [✅️️ 设置浏览器路径](#️️-设置浏览器路径)
* [✅️️ 设置用户数据路径](#️️-设置用户数据路径)
* [✅️️ 复制默认 ini 文件到当前路径](#️️-复制默认-ini-文件到当前路径)
* [✅️️ 启动浏览器](#️️-启动浏览器)

* 🧰 进阶使用
* ⚙️ 与其它项目对接

本页总览

⚙️ 与其它项目对接
==========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

DrissionPage 提供 2 个小工具，用于与 selenium 和 playwright 项目对接。

可从旧项目对象中生成`ChromiumPage`对象。

注意

只支持 chromium 内核的浏览器。

✅️️ 与 selenium 对接[​](#️️-与-selenium-对接 "✅️️ 与 selenium 对接的直接链接")
----------------------------------------------------------------

`from_selenium()`方法接收 selenium 的`WebDriver`对象，返回`ChromiumPage`对象。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `driver` | `WebDriver` | 必填 | selenium 的`WebDriver`对象 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumPage` | `ChromiumPage`对象 |

```
from DrissionPage.common import from_selenium  
from selenium.webdriver import Chrome  
  
# 创建WebDriver对象  
driver = Chrome()  
  
# 从该WebDriver对象创建ChromiumPage对象  
page = from_selenium(driver)  
  
# 用ChromiumPage对象操作浏览器  
page.get('http://DrissionPage.cn')
```

---

✅️️ 与 playwright 对接[​](#️️-与-playwright-对接 "✅️️ 与 playwright 对接的直接链接")
----------------------------------------------------------------------

`from_playwright()`方法接收 playwright 的`Page`或`Browser`对象，返回`ChromiumPage`对象。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `page_or_browser` | `Page` `Browser` | 必填 | playwright 的`Page`或`Browser`对象 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumPage` | `ChromiumPage`对象 |

```
from DrissionPage.common import from_playwright  
from playwright.sync_api import sync_playwright  
  
with sync_playwright() as p:  
    browser = p.chromium.launch()  # 用playwright启动浏览器  
    pw_page = browser.new_page()  # 创建一个新的页面  
      
    # 从Page对象创建ChromiumPage对象  
    page = from_playwright(pw_page)  
    # 或 从Browser对象创建ChromiumPage对象  
    page = from_playwright(browser)  
      
    # 用ChromiumPage对象操作浏览器  
    page.get("http://DrissionPage.cn")
```

[上一页

⚙️ 实用工具](/advance/tools)

* [✅️️ 与 selenium 对接](#️️-与-selenium-对接)
* [✅️️ 与 playwright 对接](#️️-与-playwright-对接)

* 🧰 进阶使用
* ⚙️ 异常的使用

本页总览

⚙️ 异常的使用
========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

本节介绍 DrissionPage 中的自定义异常。

✅️️ 导入[​](#️️-导入 "✅️️ 导入的直接链接")
-------------------------------

各种异常放在`DrissionPage.errors`路径中。

```
from DrissionPage.errors import *
```

✅️️ 异常介绍[​](#️️-异常介绍 "✅️️ 异常介绍的直接链接")
-------------------------------------

### 📌 `ElementNotFoundError`[​](#-elementnotfounderror "-elementnotfounderror的直接链接")

找不到元素时抛出。

---

### 📌 `AlertExistsError`[​](#-alertexistserror "-alertexistserror的直接链接")

执行 JS 或调用通过 JS 实现的功能时，若存在未处理的弹出框则抛出。

---

### 📌 `ContextLostError`[​](#-contextlosterror "-contextlosterror的直接链接")

页面被刷新后仍调用其中的元素时抛出。

---

### 📌 `ElementLostError`[​](#-elementlosterror "-elementlosterror的直接链接")

元素因页面或自身被刷新而失效后，仍对其进行调用时抛出。

---

### 📌 `CDPError`[​](#-cdperror "-cdperror的直接链接")

调用 cdp 方法产生异常时抛出。

---

### 📌 `PageDisconnectedError`[​](#-pagedisconnectederror "-pagedisconnectederror的直接链接")

页面关闭或连接断开后仍调用其功能时抛出。

---

### 📌 `JavaScriptError`[​](#-javascripterror "-javascripterror的直接链接")

JavaScript 运行错误时抛出。

---

### 📌 `NoRectError`[​](#-norecterror "-norecterror的直接链接")

对没有大小和位置信息的元素获取这些信息时抛出。

---

### 📌 `BrowserConnectError`[​](#-browserconnecterror "-browserconnecterror的直接链接")

连接浏览器出错时抛出。

---

### 📌 `NoResourceError`[​](#-noresourceerror "-noresourceerror的直接链接")

浏览器元素`src()`和`save()`获取资源失败时抛出。

---

### 📌 `CanNotClickError`[​](#-cannotclickerror "-cannotclickerror的直接链接")

---

点击元素时如元素不可点击，且设置允许抛出时抛出。

### 📌 `GetDocumentError`[​](#-getdocumenterror "-getdocumenterror的直接链接")

获取页面文档失败时抛出

---

获取页面文档失败时抛出。

### 📌 `WaitTimeoutError`[​](#-waittimeouterror "-waittimeouterror的直接链接")

自动等待失败，且设置允许抛出时抛出。

---

### 📌 `IncorrectURLError`[​](#-incorrecturlerror "-incorrecturlerror的直接链接")

访问格式不正确的 url 时抛出。

---

### 📌 `StorageError`[​](#-storageerror "-storageerror的直接链接")

操作数据时，如网站禁止操作则抛出。

---

### 📌 `CookieFormatError`[​](#-cookieformaterror "-cookieformaterror的直接链接")

导入 cookie 时如格式不正确则抛出。

---

### 📌 `LocatorError`[​](#-locatorerror "-locatorerror的直接链接")

传入的定位符格式不正确时抛出。

---

### 📌 `UnknownError`[​](#-unknownerror "-unknownerror的直接链接")

发生未知错误时抛出。

[上一页

⚙️ 命令行的使用](/advance/commands)[下一页

⚙️ 数据读取加速](/advance/accelerate)

* [✅️️ 导入](#️️-导入)
* [✅️️ 异常介绍](#️️-异常介绍)
  + [📌 `ElementNotFoundError`](#-elementnotfounderror)
  + [📌 `AlertExistsError`](#-alertexistserror)
  + [📌 `ContextLostError`](#-contextlosterror)
  + [📌 `ElementLostError`](#-elementlosterror)
  + [📌 `CDPError`](#-cdperror)
  + [📌 `PageDisconnectedError`](#-pagedisconnectederror)
  + [📌 `JavaScriptError`](#-javascripterror)
  + [📌 `NoRectError`](#-norecterror)
  + [📌 `BrowserConnectError`](#-browserconnecterror)
  + [📌 `NoResourceError`](#-noresourceerror)
  + [📌 `CanNotClickError`](#-cannotclickerror)
  + [📌 `GetDocumentError`](#-getdocumenterror)
  + [📌 `WaitTimeoutError`](#-waittimeouterror)
  + [📌 `IncorrectURLError`](#-incorrecturlerror)
  + [📌 `StorageError`](#-storageerror)
  + [📌 `CookieFormatError`](#-cookieformaterror)
  + [📌 `LocatorError`](#-locatorerror)
  + [📌 `UnknownError`](#-unknownerror)

* 🧰 进阶使用
* ⚙️ 配置文件的使用

本页总览

⚙️ 配置文件的使用
==========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

本库使用 ini 文件记录浏览器或`Session`对象的启动配置。便于配置复用，免于在代码中加入繁琐的配置信息。  
默认情况下，页面对象启动时自动加载文件中的配置信息。  
也可以在默认配置基础上用简单的方法再次修改，再保存到 ini 文件。  
也可以保存多个 ini 文件，按不同项目需要调用。

注意

* ini 文件仅用于管理启动配置，页面对象创建后再修改 ini 文件是没用的。
* 如果是接管已打开的浏览器，这些设置也没有用。
* 每次升级本库，ini 文件都会被重置，可另存到其它路径以免重置。

✅️️ ini 文件内容[​](#️️-ini-文件内容 "✅️️ ini 文件内容的直接链接")
-------------------------------------------------

ini 文件初始内容如下。

```
[paths]  
download_path =   
tmp_path =   
  
[chromium_options]  
address = 127.0.0.1:9222  
browser_path = chrome  
arguments = ['--no-default-browser-check', '--disable-suggestions-ui', '--no-first-run', '--disable-infobars', '--disable-popup-blocking', '--hide-crash-restore-bubble', '--disable-features=PrivacySandboxSettings4']  
extensions = []  
prefs = {'profile.default_content_settings.popups': 0, 'profile.default_content_setting_values': {'notifications': 2}}  
flags = {}  
load_mode = normal  
user = Default  
auto_port = False  
system_user_path = False  
existing_only = False  
new_env = False  
  
[session_options]  
headers = {'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_6) AppleWebKit/603.3.8 (KHTML, like Gecko) Version/10.1.2 Safari/603.3.8', 'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8', 'connection': 'keep-alive', 'accept-charset': 'GB2312,utf-8;q=0.7,*;q=0.7'}  
  
[timeouts]  
base = 10  
page_load = 30  
script = 30  
  
[proxies]  
http =  
https =   
  
[others]  
retry_times = 3  
retry_interval = 2
```

---

✅️️ 文件位置[​](#️️-文件位置 "✅️️ 文件位置的直接链接")
-------------------------------------

默认配置文件存放在 DrissionPage 库 `'_configs'` 文件夹下，文件名为 configs.ini。  
用户可另存其它配置文件，或从另存的文件读取配置，但默认文件的位置和名称不会改变。

---

✅️️ 使用默认配置文件启动[​](#️️-使用默认配置文件启动 "✅️️ 使用默认配置文件启动的直接链接")
-------------------------------------------------------

### 📌 使用页面对象自动加载[​](#-使用页面对象自动加载 "📌 使用页面对象自动加载的直接链接")

 这是默认启动方式。

```
from DrissionPage import Chromium  
  
browser = Chromium()
```

---

### 📌 使用配置对象加载[​](#-使用配置对象加载 "📌 使用配置对象加载的直接链接")

这种方 式一般用于加载配置后需要进一步修改。

```
from DrissionPage import ChromiumOptions, SessionOptions, Chromium  
  
co = ChromiumOptions(ini_path=r'D:\setting.ini')  
so = SessionOptions(ini_path=r'D:\setting.ini')  
  
browser = Chromium(addr_or_opts=co, session_options=so)
```

---

  

✅️️ 保存/另存 ini 文件[​](#️️-保存另存-ini-文件 "✅️️ 保存/另存 ini 文件的直接链接")
------------------------------------------------------------

```
from DrissionPage import ChromiumOptions  
  
co = ChromiumOptions()  
  
# 修改一些设置  
co.no_imgs()  
  
# 保存到当前打开的ini文件  
co.save()  
# 保存到指定位置的配置文件  
co.save(r'D:\config1.ini')  
# 保存到默认配置文件  
co.save_to_default()
```

---

✅️️ 在项目路径使用 ini 文件[​](#️️-在项目路径使用-ini-文件 "✅️️ 在项目路径使用 ini 文件的直接链接")
-------------------------------------------------------------------

默认 ini 文件存放在 DrissionPage 安装目录下，修改要通过代码进行，给调试带来不便。

因此，提供了一个便捷的方法把默认 ini 文件复制到当前项目文件夹，并且程序会优先使用项目文件夹下的 ini 文件进行初始化配置。

这样开发者可方便地手动更改配置。项目打包也可以直接打包而不会造成找不到文件问题。

复制到项目下的 ini 文件名为`'dp_configs.ini'`，程序会默认读取这个文件的配置。

### 📌 `configs_to_here()`[​](#-configs_to_here "-configs_to_here的直接链接")

此方法放在 `DrissionPage.common` 路径中，用于把默认 ini 文件复制到当前路径，并命名为`'dp_configs.ini'`。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `save_name` | `str` | `None` | 指定文件名，为`None`则命名为`'dp_configs.ini'` |

**返回：**`None`

**示例：**

在项目新建一个 py 文件，输入以下内容并运行

```
from DrissionPage.common import configs_to_here  
  
configs_to_here()
```

之后，项目文件夹会多出一个`'dp_configs.ini'`文件。页面对象初始化时会优先读取这个文件。

### 📌 用命令行复制[​](#-用命令行复制 "📌 用命令行复制的直接链接")

除了用`configs_to_here()`方法复制 ini 文件到项目文件夹，还可以用命令行方式复制。

在项目文件夹路径下运行以下命令即可：

```
dp --configs-to-here
```

效果和`configs_to_here()`一致，只是不能指定文件名。

[上一页

⤵️ 浏览器下载](/download/browser)[下一页

⚙️ 全局设置](/advance/settings)

* [✅️️ ini 文件内容](#️️-ini-文件内容)
* [✅️️ 文件位置](#️️-文件位置)
* [✅️️ 使用默认配置文件启动](#️️-使用默认配置文件启动)
  + [📌 使用页面对象自动加载](#-使用页面对象自动加载)
  + [📌 使用配置对象加载](#-使用配置对象加载)
* [✅️️ 保存/另存 ini 文件](#️️-保存另存-ini-文件)
* [✅️️ 在项目路径使用 ini 文件](#️️-在项目路径使用-ini-文件)
  + [📌 `configs_to_here()`](#-configs_to_here)
  + [📌 用命令行复制](#-用命令行复制)

* 🧰 进阶使用
* ⚙️ 打包程序

本页总览

⚙️ 打包程序
=======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

本节介绍打包程序需要注意的事项。

✅️️ 使用新的虚拟环境！[​](#️️-使用新的虚拟环境 "✅️️ 使用新的虚拟环境！的直接链接")
---------------------------------------------------

养成打包的良好习惯，使用新建的虚拟环境，只安装必要的库去打包，可以使打包出来的 exe 文件体积缩小。  
在只安装 DrissionPage 的环境中打包出来的程序，大小约为 14M。

如果您打包出来的程序体积巨大，请尝试这个方法。

---

✅️️ 解决 ini 不存在报错[​](#️️-解决-ini-不存在报错 "✅️️ 解决 ini 不存在报错的直接链接")
-------------------------------------------------------------

说明

从`v4.0.4.1`开始不存在这个报错，以下是此版本之前打包报错的解决方法。

因为程序用到 ini 文件，而打包时不会自动带上，因此直接打包是会导致运行出错。

解决办法：

* 手动带上 ini 文件，并在程序中指定路径
* 把配置信息写在程序中，不带 ini 文件

### 📌 带上 ini 文件[​](#-带上-ini-文件 "📌 带上 ini 文件的直接链接")

在程序中用相对路径方式指定 ini 文件，把 ini 文件复制到程序文件夹。

```
from DrissionPage import Chromium, ChromiumOptions, SessionOptions  
  
co = ChromiumOptions(ini_path=r'.\configs.ini')  
so = SessionOptions(ini_path=r'.\configs.ini')  
browser = Chromium(addr_or_opts=co, session_options=so)
```

可以使用`configs_to_here()`方法自动复制 ini 文件。

在项目新建一个 py 文件，输入以下内容并运行

```
from DrissionPage.common import configs_to_here  
  
configs_to_here()
```

之后，项目文件夹会多出一个`'dp_configs.ini'`文件。页面对象初始化时会优先读取这个文件。

把它和打包出来的可执行文件放在一起即可。

---

### 📌 不使用 ini[​](#-不使用-ini "📌 不使用 ini的直接链接")

在程序中指定不使用 ini 文件，就不会报错。这种方法需把所有配置信息写到代码里。

```
from DrissionPage import Chromium, ChromiumOptions, SessionOptions  
  
co = ChromiumOptions(read_file=False)  # 不读取文件方式新建配置对象  
co.set_browser_path(r'.\chrome.exe')  # 输入配置信息  
so = SessionOptions(read_file=False)  
  
browser = Chromium(addr_or_opts=co, session_options=so)
```

注意，这个时候 driver 和 session 两个参数都要输入内容，如果其中一个不需要设置可以输入`False`：

```
browser = Chromium(addr_or_opts=co, session_or_options=False)
```

---

✅️️ 实用示例[​](#️️-实用示例 "✅️️ 实用示例的直接链接")
-------------------------------------

通常，我会把一个绿色浏览器和打包后的 exe 文件放在一起，程序中用相对路径指向该浏览器，这样拿到别的电脑也可以正常使用。

```
from DrissionPage import Chromium, ChromiumOptions  
  
co = (ChromiumOptions(read_file=False)  
      .set_local_port(9888)  
      .set_cache_path(r'.\Chrome\chrome.exe')  
      .set_user_data_path(r'.\Chrome\userData'))  
browser = Chromium(addr_or_opts=co, session_options=False)  
# 注意：session_or_options=False  
tab = browser.latest_tab  
  
tab.get('http://DrissionPage.cn')
```

注意以下两点，程序就会跳过读取 ini 文件：

* `ChromiumOptions()`里要设置`read_file=False`
* 如果不传入某个模式的配置（示例中为 s 模式），要在页面对象初始化时设置对应参数为`False`

[上一页

⚙️ 数据读取加速](/advance/accelerate)[下一页

⚙️ 实用工具](/advance/tools)

* [✅️️ 使用新的虚拟环境！](#️️-使用新的虚拟环境)
* [✅️️ 解决 ini 不存在报错](#️️-解决-ini-不存在报错)
  + [📌 带上 ini 文件](#-带上-ini-文件)
  + [📌 不使用 ini](#-不使用-ini)
* [✅️️ 实用示例](#️️-实用示例)

* 🧰 进阶使用
* ⚙️ 全局设置

本页总览

⚙️ 全局设置
=======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

有一些运行时的全局设置，可以控制程序某些行为。

✅️️ 使用方式[​](#️️-使用方式 "✅️️ 使用方式的直接链接")
-------------------------------------

全局设置在`DrissionPage.common`路径中。

以`set_****()`的方式对属性进行设置。

设置方法会返回`Settings`类本身，所以支持链式操作。

使用方法：

```
from DrissionPage.common import Settings  
  
Settings.set_raise_when_wait_failed(True)  # 设置等待失败时抛出异常  
Settings.set_language('en')  # 设置报错使用英文  
  
Settings.set_raise_when_wait_failed(True).set_auto_handle_alert(True)  # 链式操作
```

---

✅️️ 设置项[​](#️️-设置项 "✅️️ 设置项的直接链接")
----------------------------------

### 📌 `set_raise_when_ele_not_found()`[​](#-set_raise_when_ele_not_found "-set_raise_when_ele_not_found的直接链接")

设置找不到元素时，是否抛出异常。初始为`False`。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_off` | `bool` | `True` | `bool`表示开或关 |

**返回：**`Settings`

---

### 📌 `set_raise_when_click_failed()`[​](#-set_raise_when_click_failed "-set_raise_when_click_failed的直接链接")

设置点击失败时，是否抛出异常。初始为`False`。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_off` | `bool` | `True` | `bool`表示开或关 |

**返回：**`Settings`

---

### 📌 `set_raise_when_wait_failed()`[​](#-set_raise_when_wait_failed "-set_raise_when_wait_failed的直接链接")

设置等待失败时，是否抛出异常。初始为`False`。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_off` | `bool` | `True` | `bool`表示开或关 |

**返回：**`Settings`

---

### 📌 `set_singleton_tab_obj()`[​](#-set_singleton_tab_obj "-set_singleton_tab_obj的直接链接")

设置 Tab 对象是否使用单例模式。初始为`True`。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_off` | `bool` | `True` | `bool`表示开或关 |

**返回：**`Settings`

---

### 📌 `set_cdp_timeout()`[​](#-set_cdp_timeout "-set_cdp_timeout的直接链接")

设置 cdp 执行超时（秒），初始为`30`。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `second` | `float` | 必填 | 秒数 |

**返回：**`Settings`

---

### 📌 `set_browser_connect_timeout()`[​](#-set_browser_connect_timeout "-set_browser_connect_timeout的直接链接")

设置连接浏览器的超时时间（秒）。初始为`30`。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `second` | `float` | 必填 | 秒数 |

**返回：**`Settings`

---

### 📌 `set_auto_handle_alert()`[​](#-set_auto_handle_alert "-set_auto_handle_alert的直接链接")

全局的自动处理弹出设置。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `accept` | `bool` `None` | `True` | 为`None`时不自动处理，为`True`时自动接受，为`False`时自动取消。 |

**返回：**`Settings`

---

### 📌 `set_language()`[​](#-set_language "-set_language的直接链接")

设置报错和提示信息语言。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `code` | `str` | 必填 | 可选`'zh_cn'`、`'en'` |

**返回：**`Settings`

---

### 📌 `set_suffixes_list()`[​](#-set_suffixes_list "-set_suffixes_list的直接链接")

设置用于解析域名后缀的本地文件路径。

默认会连网获取，离线环境下使用内置文件，可对此属性赋值手动指定路径。

通常离线环境下打包使用时需要设置。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `path` | `str` `Path` | 必填 | 文件路径 |

**返回：**`Settings`

---

  

✅️️ 示例[​](#️️-示例 "✅️️ 示例的直接链接")
-------------------------------

此示例设置找不到元素时立刻抛出异常（如不设置返回`NoneElement`）。

可直接执行查看效果。

```
from DrissionPage import SessionPage  
from DrissionPage.common import Settings  
  
Settings.set_raise_when_ele_not_found(True)  
  
page = SessionPage()  
page.get('https://www.baidu.com')  
ele = page('#abcd')
```

**输出：**

```
...前面省略...  
DrissionPage.errors.ElementNotFoundError:   
没有找到元素。  
method: ele()  
args: {'locator': '#abcd'}
```

[上一页

⚙️ 配置文件的使用](/advance/ini)[下一页

⚙️ 命令行的使用](/advance/commands)

* [✅️️ 使用方式](#️️-使用方式)
* [✅️️ 设置项](#️️-设置项)
  + [📌 `set_raise_when_ele_not_found()`](#-set_raise_when_ele_not_found)
  + [📌 `set_raise_when_click_failed()`](#-set_raise_when_click_failed)
  + [📌 `set_raise_when_wait_failed()`](#-set_raise_when_wait_failed)
  + [📌 `set_singleton_tab_obj()`](#-set_singleton_tab_obj)
  + [📌 `set_cdp_timeout()`](#-set_cdp_timeout)
  + [📌 `set_browser_connect_timeout()`](#-set_browser_connect_timeout)
  + [📌 `set_auto_handle_alert()`](#-set_auto_handle_alert)
  + [📌 `set_language()`](#-set_language)
  + [📌 `set_suffixes_list()`](#-set_suffixes_list)
* [✅️️ 示例](#️️-示例)

* 🧰 进阶使用
* ⚙️ 实用工具

本页总览

⚙️ 实用工具
=======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

`DrissionPage.common`路径可导入几个小工具。

✅️️ `make_session_ele()`[​](#️️-make_session_ele "️️-make_session_ele的直接链接")
----------------------------------------------------------------------------

此方法用于获取页面对象、元素对象或 html 文本的静态版本，或以其为基准搜索元素。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `html_or_ele` | `str` `ChromiumElement` `ChromiumPage` `ChromiumTab` `WebPage` `MixTab` `ChromiumFrame` `ShdownRoot` | 必填 | html文本、元素或页面对象 |
| `loc` | `str` `Tuple[str, str]` | `None` | 定位元组或字符串，为`None`时不在下级查找，返回根元素 |
| `index` | `int` | `1` | 获取第几个元素，从`1`开始，可传入负数获取倒数第几个，`None`获取所有 |

| 返回类型 | 说明 |
| --- | --- |
| `SessionElement` | `index`为数字时返回静态 元素对象 |
| `SessionElementsList` | `index`为`None`时返回静态元素对象组成的列表 |

**示例：**

```
from DrissionPage.common import make_session_ele  
  
html = '''  
<html><body><div>abc</div></body></html>  
'''  
ele = make_session_ele(html)  
print(ele.text)
```

**输出：**

```
abc
```

---

✅️️ `get_blob()`[​](#️️-get_blob "️️-get_blob的直接链接")
----------------------------------------------------

此方法用于获取指定 blob 资源内容。

注意

* 如果资源在异域`<iframe>`元素内，必须获取该`<iframe>`元素对象，再把该对象传入才能获取到
* 本方法只能用于获取静态的资源，流媒体不可以

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `page` | `ChromiumPage` `ChromiumTab` `WebPage` `MixTab` `ChromiumFrame` | 必填 | 该资源所在的页面对象 |
| `url` | `str` | 必填 | blob 资源 url |
| `as_bytes` | `bool` | `True` | 是否以`bytes`类型返回 |

| 返回类型 | 说明 |
| --- | --- |
| `str` | `as_bytes`参数为`False`时以 base64 格式返回 |
| `bytes` | `as_bytes`参数为`True`时以字节数据返回 |

---

✅️️ `configs_to_here()`[​](#️️-configs_to_here "️️-configs_to_here的直接链接")
-------------------------------------------------------------------------

此方法用于把默认 ini 文件复制到当前路径。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `save_name` | `str` | `None` | 指定文件名，为`None`则命名为`'dp_configs.ini'` |

**返回：** `None`

---

  

✅️️ `wait_until()`[​](#️️-wait_until "️️-wait_until的直接链接")
----------------------------------------------------------

此方法用于等待传入的方法返回值不为假。超时则抛出`TimeoutError`。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `function` | `callable` | 必填 | 要执行的方法 |
| `kwargs` | `dict` | `None` | 方法参数 |
| `timeout` | `float` | `10` | 超时时间（秒） |

**返回：** `Any`

---

✅️️ `tree()`[​](#️️-tree "️️-tree的直接链接")
----------------------------------------

此方法用于打印页面或元素结构。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `ele_or_page` | 所有页面和元素对象 | 必填 | 页面或元素对象 |
| `text` | `bool` | `False` | 是否打印元素文本 |
| `show_js` | `bool` | `False` | 打印文本时是否打印`<script>`标签内容 |
| `show_css` | `bool` | `False` | 打印文本时是否打印`<style>`标签内容 |

**返回：** `None`

---

✅️️ `Keys`[​](#️️-keys "️️-keys的直接链接")
--------------------------------------

这是快速获取特殊按键和组合键的类。

---

✅️️ `By`[​](#️️-by "️️-by的直接链接")
--------------------------------

与 selenium 的`By`类一致，方便项目迁移。

[上一页

⚙️ 打包程序](/advance/packaging)[下一页

⚙️ 与其它项目对接](/advance/docking)

* [✅️️ `make_session_ele()`](#️️-make_session_ele)
* [✅️️ `get_blob()`](#️️-get_blob)
* [✅️️ `configs_to_here()`](#️️-configs_to_here)
* [✅️️ `wait_until()`](#️️-wait_until)
* [✅️️ `tree()`](#️️-tree)
* [✅️️ `Keys`](#️️-keys)
* [✅️️ `By`](#️️-by)

![](/assets/images/color_logo-f48f02de92818fdb520db13d5406570a.png)

✨️ 概述[​](#️-概述 "✨️ 概述的直接链接")
----------------------------

DrissionPage® 是一个基于 Python 的网页自动化工具。

既能控制浏览器，也能收发数据包，还能把两者合而为一。

可兼顾浏览器自动化的便利性和 requests 的高效率。

功能强大，语法简洁优雅，代码量少，对新手友好。

---

项目地址：[gitee](https://gitee.com/g1879/DrissionPage) | [github](https://github.com/g1879/DrissionPage) | [gitcode](https://gitcode.com/g1879/DrissionPage) ⭐️ 您的星星对我非常重要！

**交流 QQ 群：** 391178600

最新版本：4.1.1.2

[![](/img/ad.png)](https://b23.tv/w62Tiqd)


---

📖 实用教程[​](#-实用教程 "📖 实用教程的直接链接")
-------------------------------

* 知识星球：[点击查看](https://wx.zsxq.com/group/48888855245858)
* Bilibili：[点击查看](https://space.bilibili.com/20526000)
* 小红书：[点击查看](https://www.xiaohongshu.com/user/profile/63293b09000000002303e82e)
* 离线文档：[点击查看](https://mall.bilibili.com/neul-next/detailuniversal/detail.html?isMerchant=1&page=detailuniversal_detail&saleType=10&itemsId=11372029&loadingShow=1&noTitleBar=1&msource=merchant_share)

![](http://drissionpage.cn/codes.png)

---

☕ 打赏作者[​](#-打赏作者 "☕ 打赏作者的直接链接")
-------------------------------

作者是个人开发者，开发和写文档工作量繁重。

如果本项目对您有所帮助，不妨打赏一下 :)

![](/assets/images/code-a284f77fdce390108ed6e3c38fbe2995.jpg)

---

💡 理念和含义[​](#-理念和含义 "💡 理念和含义的直接链接")
----------------------------------

理念：简洁而强大！

Drission 是本库作者自创的单词，为 Driver 和 Session 的合体。

因此 Drission 读作 “拽神”，Page 则表示以页面为单位操作。

---

📝 使  用条款[​](#-使用条款 "📝 使用条款的直接链接")
---------------------------------

* 禁止将 DrissionPage 应用到任何可能违反当地法律规定和道德约束的项目中
* 禁止将 DrissionPage 用于任何可能有损他人利益的项目中
* 禁止将 DrissionPage 用于攻击与骚扰行为
* 遵守 Robots 协议，禁止将 DrissionPage 用于采集法律或系统 Robots 协议不允许的数据

使用 DrissionPage 发生的一切行为均由使用人自行负责。
因使用 DrissionPage 进行任何行为所产生的一切纠纷及后果均与版权持有人无关，
版权持有人不承担任何使用 DrissionPage 带来的风险和损失。
版权持有人不对 DrissionPage 可能存在的缺陷导致的任何损失负任何责任。

个人或组织如未获得版权持有人授权，不得将本项目以源代码或二进制形式用于商业行为。

DrissionPage 源代码和本文档内容，未获得作者授权禁止用于出版。

* [✨️ 概述](#️-概述)
* [📖 实用教程](#-实用教程)
* [☕ 打赏作者](#-打赏作者)
* [💡 理念和含义](#-理念和含义)
* [📝 使用条款](#-使用条款)

Markdown page example
=====================

You don't need React to write simple standalone pages.

在文档中搜索 | DrissionPage官网




[跳到主要内容](#__docusaurus_skipToContent_fallback)

支持开源作者，请关闭广告屏蔽功能。当前文档适用于：DrissionPage **4.1.1.2**

[![DrissionPage](/img/color_logo.png)![DrissionPage](/img/color_logo.png)

**DrissionPage**](/)[特性](/features/4.1)[入门](/get_start/installation)[文档](/browser_control/intro)[教程](/tutorials/xingqiu)[进度](/versions/4.1.x)[支持](/support)

[更多作品](#)

* [DrissionRecord](https://drissionpage.cn/DrissionRecord)
* [DownloadKit](https://drissionpage.cn/DownloadKitDocs)
* [TimePinner](https://drissionpage.cn/TimePinnerDocs)
* [MixPage](https://drissionpage.cn/MixPageDocs)
* [3.2 版文档](https://mall.bilibili.com/neul-next/detailuniversal/detail.html?isMerchant=1&page=detailuniversal_detail&saleType=10&itemsId=12019346&loadingShow=1&noTitleBar=1&msource=merchant_share)
* [4.0 版文档](https://mall.bilibili.com/neul-next/detailuniversal/detail.html?isMerchant=1&page=detailuniversal_detail&saleType=10&itemsId=12020073&loadingShow=1&noTitleBar=1&msource=merchant_share)

[项目地址](#)

* [Gitee](https://gitee.com/g1879/DrissionPage)
* [GitHub](https://github.com/g1879/DrissionPage)
* [GitCode](https://gitcode.com/g1879/DrissionPage)

搜索

在文档中搜索
======

作者

* [g1879](https://gitee.com/g1879)

交流

* [联系邮箱：g1879@qq.com](#)
* [QQ群：391178600](#)

旧版地址

* [4.0 版文档](https://mall.bilibili.com/neul-next/detailuniversal/detail.html?isMerchant=1&page=detailuniversal_detail&saleType=10&itemsId=12020073&loadingShow=1&noTitleBar=1&msource=merchant_share)
* [3.2 版文档](https://mall.bilibili.com/neul-next/detailuniversal/detail.html?isMerchant=1&page=detailuniversal_detail&saleType=10&itemsId=12019346&loadingShow=1&noTitleBar=1&msource=merchant_share)
* [MixPage](https://DrissionPage.cn/mixpagedocs)

本文档禁止商用 [DrissionPageDocs](https://drissionpage.cn) by g1879 is licensed under [CC BY-NC 4.0](http://creativecommons.org/licenses/by-nc/4.0/?ref=chooser-v1)

DrissionPage®为作者已注册的商标　　[粤ICP备2024179482号-1](https://beian.miit.gov.cn/).

本页总览

支持作者
====

  

✅️️ 关于作者[​](#️️-关于作者 "✅️️ 关于作者的直接链接")
-------------------------------------

DrissionPage 作者：g1879。

编程爱好者。

联系邮箱：[g1879@qq.com](mailto:g1879@qq.com)

作者 QQ：3970203862。

---

✅️️ 打赏作者[​](#️️-打赏作者 "✅️️ 打 赏作者的直接链接")
--------------------------------------

作者是个人开发者，开发和写文档工作量繁重。

如果本项目对您有所帮助，不妨打赏一下 :)

![](/assets/images/code1-5d7b571a5c225a47606640fc02c0d6ba.jpg)

---

✅️️ 作者小店[​](#️️-作者小店 "✅️️ 作者小店的直接链接")
-------------------------------------

作者 B 站小店可购买离线文档、浏览器插件、实战代码等。

![](/assets/images/barrack-d92b34e66fbfb806f484614a6b8f8833.png)

---

✅️️ 商用授权[​](#️️-商用授权 "✅️️ 商用授权的直接链接")
-------------------------------------

DrissionPage 可个人免费使用，须商用请联系作者获取授权。

---

✅️️ 网站广告[​](#️️-网站广告 "✅️️ 网站广告的直接链接")
-------------------------------------

承接本站广告，发邮件或 QQ 详谈。

---

✅️️ 技术咨询和接单[​](#️️-技术咨询和接单 "✅️️ 技术咨询和接单的直接链接")
----------------------------------------------

发邮件或 QQ 详谈。

* [✅️️ 关于作者](#️️-关于作者)
* [✅️️ 打赏作者](#️️-打赏作者)
* [✅️️ 作者小店](#️️-作者小店)
* [✅️️ 商用授权](#️️-商用授权)
* [✅️️ 网站广告](#️️-网站广告)
* [✅️️ 技术咨询和接单](#️️-技术咨询和接单)

## <a name="浏览器控制"></a>浏览器控制

* 🚀 控制浏览器
* 🛰️ 动作链

本页总览

🛰️ 动作链
======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

动作链可以在浏览器上完成一系列交互行为，如鼠标移动、鼠标点击、键盘输入等。

浏览器页面对象都支持使用动作链。

可以链式操作，也可以分开执行，每个动作执行即生效，无需`perform()`。

这些操作皆为模拟，真正的鼠标不会移动，因此可以多个标签页同时操作。

✅️ 使用方法[​](#️-使用方法 "✅️ 使用方法的直接链接")
----------------------------------

可以用上述对象内置的`actions`属性调用动作链，也可以主动创建一个动作链对象，将页面对象传入使用。

这两种方式唯一区别是，前者会等待页面加载完毕再执行，后者不会。

### 📌 使用内置`actions`属性[​](#-使用内置actions属性 "-使用内置actions属性的直接链接")

说明

这种方式会等到页面框架文档（不包括 js 数据）加载完成再执行动作。

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.get('https://www.baidu.com')  
tab.actions.move_to('#kw').click().type('DrissionPage')  
tab.actions.move_to('#su').click()
```

---

### 📌 使用新对象[​](#-使用新对象 "📌 使用新对象的直接链接")

使用`from DrissionPage.common import Actions`导入动作链。

只要把`WebPage`对象或`ChromiumPage`对象传入即可。动作链只在这个页面上生效。

| 初始化参数 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `page` | `ChromiumPage` `WebPage` `ChromiumTab` | 必填 | 动作链要操作的浏览器页面 |

说明

这种方式**不会**等到页面框架文档（不包括 js 数据）加载完成再执行动作。

**示例：**

```
from DrissionPage import Chromium  
from DrissionPage.common import Actions  
  
tab = Chromium().latest_tab  
ac = Actions(tab)  
tab.get('https://www.baidu.com')  
ac.move_to('#kw').click().type('DrissionPage')  
ac.move_to('#su').click()
```

---

### 📌 操作方式[​](#-操作方式 "📌 操作方式的直接链接")

多个动作可以用链式模式操作：

```
tab.actions.move_to(ele).click().type('some text')
```

也可以多个操作分开执行：

```
tab.actions.move_to(ele)  
tab.actions.click()  
tab.actions.type('some text')
```

这两种方式效果是一样的，每个动作总会依次执行。

---

✅️ 移动鼠标[​](#️-移动鼠标 "✅️ 移动鼠标的直接链接")
----------------------------------

### 📌 `move_to()`[​](#-move_to "-move_to的直接链接")

此方法用于移动鼠标到元素中点，或页面  上的某个绝对坐标。

当`offset_x`和`offset_y`都为`None`时，移动到元素中间点。

当传入偏移量时，偏移量相对于元素左上角坐标。

| 初始化参数 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `ele_or_loc` | `ChrmoiumElement` `str` `Tuple[int, int]` | 必填 | 元素对象、文本定位符或绝对坐标，坐标为`tuple`(int, int) 形式 |
| `offset_x` | `int` | `None` | x 轴偏移量，向右为正，向左为负 |
| `offset_y` | `int` | `None` | y 轴偏移量，向下为正，向上为负 |
| `duration` | `float` | `0.5` | 拖动用时，传入`0`即瞬间到达 |

| 返回类型 | 说明 |
| --- | --- |
| `Actions` | 动作链对象本身 |

**示例：** 使鼠标移动到 ele 元素上

```
ele = tab('tag:a')  
tab.actions.move_to(ele_or_loc=ele)
```

---

### 📌 `move()`[​](#-move "-move的直接链接")

此方法用于使鼠标相对当前位置移动若干距离。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `offset_x` | `int` | `0` | x 轴偏移量，向右为正，向左为负 |
| `offset_y` | `int` | `0` | y 轴偏移量，向下为正，向上为负 |
| `duration` | `float` | `0.5` | 拖动用时，传入`0`即瞬间到达 |

| 返回类型 | 说明 |
| --- | --- |
| `Actions` | 动作链对象本身 |

**示例：** 鼠标向右移动 300 像素

```
tab.actions.move(300, 0)
```

---

### 📌 `up()`[​](#-up "-up的直接链接")

此方法用于使鼠标相对当前位置向上移动若干距离。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `pixel` | `int` | 必填 | 鼠标移动的像素值 |

| 返回类型 | 说明 |
| --- | --- |
| `Actions` | 动作链对象本身 |

**示例：** 鼠标向上移动 50 像素

```
tab.actions.up(50)
```

---

### 📌 `down()`[​](#-down "-down的直接链接")

此方法用于使鼠标相对当前位置向下移动若干距离。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `pixel` | `int` | 必填 | 鼠标移动的像素值 |

| 返回类型 | 说明 |
| --- | --- |
| `Actions` | 动作链对象本身 |

**示例：**

```
tab.actions.down(50)
```

---

### 📌 `left()`[​](#-left "-left的直接链接")

此方法用于使鼠标相对当前位置向左移动若干距离。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `pixel` | `int` | 必填 | 鼠标移动的像素值 |

| 返回类型 | 说明 |
| --- | --- |
| `Actions` | 动作链对象本身 |

**示例：**

```
tab.actions.left(50)
```

---

### 📌 `right()`[​](#-right "-right的直接链接")

此方法用于使鼠标相对当前位置向右移动若干距离。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `pixel` | `int` | 必填 | 鼠标移动的像素值 |

| 返回类型 | 说明 |
| --- | --- |
| `Actions` | 动作链对象本身 |

**示例：**

```
tab.actions.right(50)
```

---

✅️ 鼠标按键[​](#️-鼠标按键 "✅️ 鼠标按键的直接链接")
----------------------------------

### 📌 `click()`[​](#-click "-click的直接链接")

此方法用于单击鼠标左键，单击前可先移动到元素上。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_ele` | `ChromiumElement` `str` | `None` | 要点击的元素对象或文本定位符 |
| `times` | `int` | `1` | 点击次数 |

| 返回类型 | 说明 |
| --- | --- |
| `Actions` | 动作链对象本身 |

**示例：**

```
tab.actions.click('#div1')
```

---

### 📌 `r_click()`[​](#-r_click "-r_click的直接链接")

此方法用于单击鼠标右键，单击前可先移动到元素上。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_ele` | `ChromiumElement` `str` | `None` | 要点击的元素对象或文本定位符 |
| `times` | `int` | `1` | 点击次数 |

| 返回类型 | 说明 |
| --- | --- |
| `Actions` | 动作链对象本身 |

**示例：**

```
tab.actions.r_click('#div1')
```

---

### 📌 `m_click()`[​](#-m_click "-m_click的直接链接")

此方法用于单击鼠标中键，单击前可先移动到元素上。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_ele` | `ChromiumElement` `str` | `None` | 要点击的元素对象或文本定位符 |
| `times` | `int` | `1` | 点击次数 |

| 返回类型 | 说明 |
| --- | --- |
| `Actions` | 动作链对象本身 |

**示例：**

```
tab.actions.m_click('#div1')
```

---

### 📌 `hold()`[​](#-hold "-hold的直接链接")

此方法用于按住鼠标左键不放，按住前可先移动到元素上。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_ele` | `ChromiumElement` `str` | `None` | 要按住的元素对象或文本定位符 |

| 返回类型 | 说明 |
| --- | --- |
| `Actions` | 动作链对象本身 |

**示例：**

```
tab.actions.hold('#div1')
```

---

### 📌 `release()`[​](#-release "-release的直接链接")

此方法用于释放鼠标左键，释放前可先移动到元素上。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_ele` | `ChromiumElement` `str` | `None` | 要释放的元素对象或文本定位符 |

| 返回类型 | 说明 |
| --- | --- |
| `Actions` | 动作链对象本身 |

**示例：** 移动到某元素上然后释放鼠标左键

```
tab.actions.release('#div1')
```

---

### 📌 `r_hold()`[​](#-r_hold "-r_hold的直接链接")

此方法用于按住鼠标右键不放，按住前可先移动到元素上。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_ele` | `ChromiumElement` `str` | `None` | 要按住的元素对象或文本定位符 |

| 返回类型 | 说明 |
| --- | --- |
| `Actions` | 动作链对象本身 |

---

### 📌 `r_release()`[​](#-r_release "-r_release的直接链接")

此方法用于释放鼠标右键，释放前可先移动到元素上。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_ele` | `ChromiumElement` `str` | `None` | 要释放的元素对象或文本定位符 |

| 返回类型 | 说明 |
| --- | --- |
| `Actions` | 动作链对象本身 |

---

### 📌 `m_hold()`[​](#-m_hold "-m_hold的直接链接")

此方法用于按住鼠标中键不放，按住前可先移动到元素上。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_ele` | `ChromiumElement` `str` | `None` | 要按住的元素对象或文本定位符 |

| 返回类型 | 说明 |
| --- | --- |
| `Actions` | 动作链对象本身 |

---

### 📌 `m_release()`[​](#-m_release "-m_release的直接链接")

此方法用于释放鼠标中键，释放前可先移动到元素上。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_ele` | `ChromiumElement` `str` | `None` | 要释放的元素对象或文本定位符 |

| 类型 | 说明 |
| --- | --- |
| `Actions` | 动作链对象本身 |

---

  

✅️ 滚动滚轮[​](#️-滚动滚轮 "✅️ 滚动滚轮的直接链接")
----------------------------------

### 📌 `scroll()`[​](#-scroll "-scroll的直接链接")

此方法用于滚动鼠标滚轮，滚动前可先移动到元素上。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `delta_y` | `int` | `0` | 滚轮 y 轴变化值，向下为正，向上为负 |
| `delta_x` | `int` | `0` | 滚轮 x 轴变化值，向右为正，向左为负 |
| `on_ele` | `ChromiumElement` `str` | `None` | 要滚动的元素对象或文本定位符 |

| 返回类型 | 说明 |
| --- | --- |
| `Actions` | 动作链对象本身 |

---

✅️ 键盘按键和文本输入[​](#️-键盘按键和文本输入 "✅️ 键盘按键和文本输入的直接链接")
-------------------------------------------------

### 📌 `key_down()`[​](#-key_down "-key_down的直接链接")

此方法用于按下键盘按键。非字符串按键（如 ENTER）可输入其名称，也可以用 Keys 类获取。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `key` | `str` | 必填 | 按键名称，或从`Keys`类获取的键值 |

| 返回类型 | 说明 |
| --- | --- |
| `Actions` | 动作链对象本身 |

**示例：** 按下 ENTER 键

```
from DrissionPage.common import Keys  
  
tab.actions.key_down('ENTER')  # 输入按键名称  
tab.actions.key_down(Keys.ENTER)  # 从Keys获取按键
```

---

### 📌 `key_up()`[​](#-key_up "-key_up的直接链接")

此方法用于提起键盘按键。非字符串按键（如 ENTER）可输入其名称，也可以用 Keys 类获取。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `key` | `str` | 必填 | 按键名称，或从`Keys`类获取的键值 |

| 返回类型 | 说明 |
| --- | --- |
| `Actions` | 动作链对象本身 |

**示例：** 提起 ENTER 键

```
from DrissionPage.common import Keys  
  
tab.actions.key_up('ENTER')  # 输入按键名称  
tab.actions.key_up(Keys.ENTER)  # 从Keys获取按键
```

---

### 📌 `input()`[​](#-input "-input的直接链接")

此方法用于输入一段文本或多段文本，也可输入组合键。

多段文本或组合键用列表传入。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `text` | `str` `list` `tuple` | 必填 | 要输入的文本或按键，多段文本或组合键可用`list`或`tuple`传入 |

| 返回类型 | 说明 |
| --- | --- |
| `Actions` | 动作链对象本身 |

**示例：**

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.get('https://www.baidu.com')  
tab.actions.click('#kw').input('DrissionPage')
```

---

### 📌 `type()`[​](#-type "-type的直接链接")

此方法用于以按键盘的方式输入一段或多段文本。也可输入组合键。

`type()`与`input()`区别在于前者模拟按键输入，逐个字符按下和提起，后者直接输入一整段文本。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `keys` | `str` `list` `tuple` | 必填 | 要输入的文本或按键，多段文本或组合键可用`list`或`tuple`传入 |

| 返回类型 | 说明 |
| --- | --- |
| `Actions` | 动作链对象本身 |

**示例：**

```
# 键入一段文本  
tab.actions.type('text')  
  
# 键入多段文本  
tab.actions.type(('ab', 'cd'))  
  
# 光标向左移动一位再键入文本  
tab.actions.type((Keys.LEFT, 'abc'))  
  
# 输入快捷键  
tab.actions.type(Keys.CTRL_A)
```

---

✅️ 拖入文件和文本[​](#️-拖入文件和文本 "✅️ 拖入文件和文本的直接链接")
-------------------------------------------

### 📌 `drag_in()`[​](#-drag_in "-drag_in的直接链接")

此方法用于模拟从浏览器外部拖入文件或文本。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `ele_or_loc` | `str` `ChromiumElement` | 必填 | 接收拖动动作的元素定位符 |
| `files` | `str` `list` `tuple` | `None` | 要拖入文件路径，可多个，不为`None`时下面参数无效 |
| `text` | `str` | `None` | 要拖入的文本，`files`参数为`None`时才生效 |
| `title` | `str` | `None` | 如果`text`是超链接，可在此设置`title`，与`baseURL`互斥 |
| `baseURL` | `str` | `None` | 如果`text`是`html`，可在此设置`baseUrl`，与`title`互斥 |

| 返回类型 | 说明 |
| --- | --- |
| `Actions` | 动作链对象本身 |

---

✅️ 等待[​](#️-等待 "✅️ 等待的直接链接")
----------------------------

### 📌 `wait()`[​](#-wait "-wait的直接链接")

此方法用于等待若干秒。

`scope`为`None`时，效果与`time.sleep()`没有区别，等待指定秒数。

`scope`不为`None`时，获取两个参数之间的一个随机值，等待这个数值的秒数。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `second` | `float` | 必填 | 要等待的秒数，`scope`不为`None`时表示随机数范围起始值 |
| `scope` | `float` | `None` | 随机数范围结束值 |

**返回：**`None`

---

✅️ 属性[​](#️-属性 "✅️ 属性的直接链接")
----------------------------

### 📌 `owner`[​](#-owner "-owner的直接链接")

此属性返回使用此动作链的页面对象。

**类型：**`ChromiumBase`

---

### 📌 `curr_x`[​](#-curr_x "-curr_x的直接链接")

此属性返回当前光标位置的 x 坐标。

**类型：**`int`

---

### 📌 `curr_y`[​](#-curr_y "-curr_y的直接链接")

此属性返回当前光标位置的 y 坐标。

**类型：**`int`

---

✅️ 示例[​](#️-示例 "✅️ 示例的直接链接")
----------------------------

### 📌 模拟输入 ctrl+a[​](#-模拟输入-ctrla "📌 模拟输入 ctrl+a的直接链接")

```
from DrissionPage import Chromium  
from DrissionPage.common import Keys  
  
# 创建页面对象  
tab = Chromium().latest_tab  
  
# 鼠标移动到<input>元素上  
tab.actions.move_to('tag:input')  
# 点击鼠标，使光标落到元素中  
tab.actions.click()  
# 按下 ctrl 键  
tab.actions.key_down(Keys.CTRL)  
# 输入 a  
tab.actions.type('a')  
# 提起 ctrl 键  
tab.actions.key_up(Keys.CTRL)
```

链式写法：

```
tab.actions.click('tag:input').key_down(Keys.CTRL).type('a').key_up(Keys.CTRL)
```

更简单的写法：

```
tab.actions.click('tag:input').type(Keys.CTRL_A)
```

---

### 📌 拖拽元素[​](#-拖拽元素 "📌 拖拽元素的直接链接")

把一个元素向右拖拽 300 像素：

```
from DrissionPage import Chromium  
  
# 创建页面  
tab = Chromium().latest_tab  
  
# 左键按住元素  
tab.actions.hold('#div1')  
# 向右移动鼠标300像素  
tab.actions.right(300)  
# 释放左键  
tab.actions.release()
```

把一个元素拖拽到另一个元素上：

```
tab.actions.hold('#div1').release('#div2')
```

[上一页

🛰️ iframe 操作](/browser_control/iframe)[下一页

🛰️ 模式切换](/browser_control/mode_change)

* [✅️ 使用方法](#️-使用方法)
  + [📌 使用内置`actions`属性](#-使用内置actions属性)
  + [📌 使用新对象](#-使用新对象)
  + [📌 操作方式](#-操作方式)
* [✅️ 移动鼠标](#️-移动鼠标)
  + [📌 `move_to()`](#-move_to)
  + [📌 `move()`](#-move)
  + [📌 `up()`](#-up)
  + [📌 `down()`](#-down)
  + [📌 `left()`](#-left)
  + [📌 `right()`](#-right)
* [✅️ 鼠标按键](#️-鼠标按键)
  + [📌 `click()`](#-click)
  + [📌 `r_click()`](#-r_click)
  + [📌 `m_click()`](#-m_click)
  + [📌 `hold()`](#-hold)
  + [📌 `release()`](#-release)
  + [📌 `r_hold()`](#-r_hold)
  + [📌 `r_release()`](#-r_release)
  + [📌 `m_hold()`](#-m_hold)
  + [📌 `m_release()`](#-m_release)
* [✅️ 滚动滚轮](#️-滚动滚轮)
  + [📌 `scroll()`](#-scroll)
* [✅️ 键盘按键和文本输入](#️-键盘按键和文本输入)
  + [📌 `key_down()`](#-key_down)
  + [📌 `key_up()`](#-key_up)
  + [📌 `input()`](#-input)
  + [📌 `type()`](#-type)
* [✅️ 拖入文件和文本](#️-拖入文件和文本)
  + [📌 `drag_in()`](#-drag_in)
* [✅️ 等待](#️-等待)
  + [📌 `wait()`](#-wait)
* [✅️ 属性](#️-属性)
  + [📌 `owner`](#-owner)
  + [📌 `curr_x`](#-curr_x)
  + [📌 `curr_y`](#-curr_y)
* [✅️ 示例](#️-示例)
  + [📌 模拟输入 ctrl+a](#-模拟输入-ctrla)
  + [📌 拖拽元素](#-拖拽元素)

* 🚀 控制浏览器
* 🛰️ 浏览器对象

本页总览

🛰️ 浏览器对象
========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

我们已经了解如何创建浏览器对象，本节介绍浏览器对象的功能。

说明

文中的 “Tab 对象” 是`ChromiumTab`和`MixTab`的统称。

✅️️ 获取标签页对象或信息[​](#️️-获取标签页对象或信息 "✅️️ 获取标签页对象或信息的直接链接")
-------------------------------------------------------

### 📌 `get_tab()`[​](#-get_tab "-get_tab的直接链接")

此方法用于获取一个标签页对象或它的 id。

`id_or_num`不为`None`时，获取`id_or_num`指定的标签页。后面几个参数无效。

`id_or_num`为`None`时，根据后面几个参数指定的条件查找标签页（与关系）。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `id_or_num` | `str` `int` | `None` | 要获取的标签页 id 或序号，序号从`1`开始，可传入负数获取倒数第几个，不是视觉排列顺序，而是激活顺序 |
| `title` | `str` | `None` | 要匹配 title 的文本，模糊匹配，为`None`则匹配所有 |
| `url` | `str` | `None` | 要匹配 url 的文本，模糊匹配，为`None`则匹配所有 |
| `tab_type` | `str` `list` `tuple` | `'page'` | 标签页类型，可用列表输入多个，如`'page'`、`'iframe'`等，为`None`则匹配所有 |
| `as_id` | `bool` | `False` | 是否返回标签页 id 而不是标签页对象 |

| 返回类型 | 说明 |
| --- | --- |
| `MixTab` | `as_id`为`False`时返回获取到的标签页对象 |
| `str` | `as_id`为`True`时返回获取到的标签页的 id |

```
from DrissionPage import Chromium  
  
browser = Chromium()  
tab = browser.get_tab()
```

---

### 📌 `get_tabs()`[​](#-get_tabs "-get_tabs的直接链接")

此方法用于获取多个符合条件的`MixTab`对象或它们的 id组成的列表。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `title` | `str` | `None` | 要匹配 title 的文本，模糊匹配，为`None`则匹配所有 |
| `url` | `str` | `None` | 要匹配 url 的文本，模糊匹配，为`None`则匹配所有 |
| `tab_type` | `str` `list` `tuple` | `'page'` | 标签页类型，可用列表输入多个，如`'page'`、`'iframe'`等，为`None`则匹配所有 |
| `as_id` | `bool` | `False` | 是否返回标签页 id 而不是标签页对象 |

| 返回类型 | 说明 |
| --- | --- |
| `List[MixTab]` | `as_id`为`False`时返回获取到的标签页对象组成的列表 |
| `List[str]` | `as_id`为`True`时返回获取到的标签页的 id 组成的列表 |

---

### 📌 `latest_tab`[​](#-latest_tab "-latest_tab的直接链接")

此属性返回最新的标签页对象或 id。

* 控制本地浏览器时，返回最后激活的标签页
* 控制远程浏览器时，返回最后创建的标签页

如果关闭单例模式，即当`Settings.singleton_tab_obj`为`False`时，返回标签页的 id。

| 返回类型 | 说明 |
| --- | --- |
| `MixTab` | 单例模式时返回标签页对象 |
| `str` | 非单例模式时返回标签页 id |

---

### 📌 `tabs_count`[​](#-tabs_count "-tabs_count的直接链接")

此属性返回标签页数量，只统计普通标签页（即`'page'`、`'webview'`类型）。

**类型：**`int`

---

### 📌 `tab_ids`[​](#-tab_ids "-tab_ids的直接链接")

此属性返回所有标签页 id 组成的列表，只统计普通标签页（即`'page'`、`'webview'`类型）。

**类型：**`List[str]`

---

✅️️ 标签页操作[​](#️️-标签页操作 "✅️️ 标签页操作的直接链接")
----------------------------------------

### 📌 `new_tab()`[​](#-new_tab "-new_tab的直接链接")

此方法用于新建标签页，并返回标签页对象。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `url` | `str` | `None` | 新标签页跳转到的网址，为`None`时新建空标签页 |
| `new_window` | `bool` | `False` | 是否在新窗口打开标签页，隐身模式下无效 |
| `background` | `bool` | `False` | 是否不激活新标签页，隐身模式和访客模式及`new_window`为`True`时无效 |
| `new_context` | `bool` | `False` | 是否创建独立环境，隐身模式和访客模式下无效 |

| 返回类型 | 说明 |
| --- | --- |
| `MixTab` | 标签页对象 |

---

### 📌 `activate_tab()`[​](#-activate_tab "-activate_tab的直接链接")

此方法用于使一个标签页显示到前端。可传入 Tab 对象、标签页 id、标签页序号。

注意标签页序号不是视觉顺序，而是激活顺序。

说明

标签页没有焦点的概念，多个标签页可以并行操作，这个方法不会对所谓焦点产生什么影响。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `id_ind_tab` | `str` `int` `ChromiumTab` `MixTab` | 必填 | 标签页 id（`str`）、Tab 对象或标签页序号（`int`），序号从`1`开始 |

**返回：**`None`

---

### 📌 `close_tabs()`[​](#-close_tabs "-close_tabs的直接链接")

此方法用于关闭标签页。可指定多个，可关闭指定标签页以外的。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `tabs_or_ids` | `str` `ChromiumTab` `MixTab` `List|Tuple[ChromiumTab|MixTab|str]` | 必填 | 指定的标签页对象或 id，可用列表或元组传入多个 |
| `others` | `bool` | `False` | 是否关闭指定标签页之外的 |

**返回：**`None`

---

### 📌 单例模式说明[​](#-单例模式说明 "📌 单例模式说明的直接链接")

默认设置下，一个标签页只有一个 Tab 对象。

对同一个标签页反复使用`get_tab()`获取到的是同一个对象。

如上文所述，`latest_tab`获取的也是曾经生成过的 Tab 对象。

如果需要多个 Tab 对象共同管理一个标签页，可关闭单例模式：

```
from DrissionPage.common import Settings  
  
Settings.set_singleton_tab_obj(False)
```

关闭后，每次`get_tab()`都会创建新的 Tab 对象，`latest_tab`改成返回 Tab 对象的 id。

```
from DrissionPage import Chromium  
from DrissionPage.common import Settings  
  
Settings.set_singleton_tab_obj(False)  
browser = Chromium()  
tab1 = browser.get_tab()  
tab2 = browser.get_tab()  
print(tab1.title, id(tab1))  
print(tab2.title, id(tab2))
```

**输出：**

```
新标签页 2500121968848  
新标签页 2500125672272
```

---

✅️️ 浏览器运行参数[​](#️️-  浏览器运行参数 "✅️️ 浏览器运行参数的直接链接")
------------------------------------------------

浏览器运行设置是一些总体的运行参数。

新标签页对象会继承浏览器的运行设置，但标签页对象后再修改浏览器设置，已生成的设置也不会改变。

设置优先级：Tab 对象设置 > `Chromium`对象设置 > `Settings`设置

### 📌 `user_data_path`[​](#-user_data_path "-user_data_path的直接链接")

此参数返回浏览器返回用户文件夹路径。

**类型：**`str`

---

### 📌 `download_path`[​](#-download_path "-download_path的直接链接")

此参数返回浏览器返回默认下载路径。

**类型：**`str`

---

### 📌 几种超时参数[​](#-几种超时参数 "📌 几种超时参数的直接链接")

此参数返回所有超时设置，单位为秒，有`base`、`page_load`、`script`三种。

* `timeouts.base`：各种等待的基础超时设置
* `timeouts.page_load`：页面文档加载的超时设置
* `timeouts.script`：JavaScript 运行超时设置

**类型：**`float`

---

### 📌 `timeout`[​](#-timeout "-timeout的直接链接")

此参数返回基础超时设置，单位为秒，即`timeouts.base`。

**类型：**`float`

---

### 📌 `load_mode`[​](#-load_mode "-load_mode的直接链接")

此参数返回页面加载模式，包括`'none'`、`'normal'`、`'eager'`三种。

**类型：**`str`

---

  

✅️️ 浏览器运行设置[​](#️️-浏览器运行设置 "✅️️ 浏览器运行设置的直接链接")
----------------------------------------------

### 📌 `set.timeouts()`[​](#-settimeouts "-settimeouts的直接链接")

此方法用于设置运行时的各种超时时间，单位为秒。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `base` | `float` | `None` | 各种等待的默认超时时间，为`None`则不修改 |
| `page_load` | `float` | `None` | 页面文档加载超时时间，为`None`则不修改 |
| `script` | `float` | `None` | 脚本运行超时时间，为`None`则不修改 |

**返回：**`None`

---

### 📌 加载模式设置[​](#-加载模式设置 "📌 加载模式设置的直接链接")

此方法用于设置页面加载模式。具体使用方法详见访问网页章节。

* `set.load_mode.normal()`：等待所有资源加载完毕的模式
* `set.load_mode.eager()`：等待文档加载完即停止加载的模式
* `set.load_mode.none()`：不会主动停止加载的模式

**返回：**`None`

---

### 📌 `set.retry_times()`[​](#-setretry_times "-setretry_times的直接链接")

此方法用于设置页面连接失败重连次数。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `times` | `int` | 必填 | 重连次数 |

**返回：**`None`

---

### 📌 `set.retry_interval()`[​](#-setretry_interval "-setretry_interval的直接链接")

此方法用于设置连接失败重连间隔（秒）。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `interval` | `float` | 必填 | 重连间隔 |

**返回：**`None`

---

### 📌 `set.cookies()`[​](#-setcookies "-setcookies的直接链接")

此方法用于设置一个或多个 cookie。

注意

用这个方法设置 cookies 记得带上`domain`属性。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `cookies` | `CookieJar` `Cookie` `list` `tuple` `str` `dict` | 必填 | 支持多种格式的 cookies 信息，一个或多个都可以 |

**返回：**`None`

---

### 📌 `set.cookies.clear()`[​](#-setcookiesclear "-setcookiesclear的直接链接")

此方法用于清除浏览器所有 cookies。

**参数：** 无

**返回：**`None`

---

### 📌 `set.auto_handle_alert()`[​](#-setauto_handle_alert "-setauto_handle_alert的直接链接")

此方法用于设置是否启用自动处理 alert 弹窗。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_off` | `bool` | `True` | `bool`表示开或关，传入`None`表示使用`Settings`设置 |
| `accept` | `bool` | `True` | 处理 alert 的方式，确定还是取消 |

**返回：**`None`

---

### 📌 `set.download_path()`[​](#-setdownload_path "-setdownload_path的直接链接")

此方法用于设置下载文件默认保存路径。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `path` | `Path` `str` `None` | 必填 | 文件夹路径，传入`None`表示当前文件夹 |

**返回：**`None`

---

### 📌 `set.download_file_name()`[​](#-setdownload_file_name "-setdownload_file_name的直接链接")

此方法用于设置下一个被下载文件的名称。

有些下载是从临时闪现的标签页触发的，这种需要由浏览器对象去捕捉和设置下载信息。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `name` | `str` | `None` | 文件名，可不含后缀，会自动使用远程文件后缀，为`None`使用远程文件名 |
| `suffix` | `str` | `None` | 后缀名，显式设置后缀名，不使用远程文件后缀 |

**返回：**`None`

---

### 📌 `set.when_download_file_exists()`[​](#-setwhen_download_file_exists "-setwhen_download_file_exists的直接链接")

此方法用于设置当存在同名文件时的处理方式。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `mode` | `str` | 必填 | 可在`'rename'`、`'overwrite'`、`'skip'`、`'r'`、`'o'`、`'s'`中选择 |

* `'rename'`或`'r'`：自动重命名，在文件名后加序号，如`'_1'`
* `'overwrit'`或`'o'`：覆盖已有文件
* `'skip'`或`'s'`：跳过，不下载

**返回：**`None`

---

### 📌 `set.NoneElement_value()`[​](#-setnoneelement_value "-setnoneelement_value的直接链接")

此方法用于设置查找元素失败时返回的空元素是否返回设定值。详见元素查找行为章节。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `value` | `Any` | `None` | 设置的设定值 |
| `on_off` | `bool` | `True` | 是否启用 |

**返回：**`None`

---

✅️️ 浏览器信息[​](#️️-浏览器信息 "✅️️ 浏览器信息的直接链接")
----------------------------------------

### 📌 `cookies()`[​](#-cookies "-cookies的直接链接")

此方法以列表形式返回浏览器所有域名的 cookies，cookie 是`dict`格式。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `all_info` | `bool` | `False` | 是否返回所有内容，`False`则只返回`'name'`、`'value'`、`'domain'`三个属性 |

**返回：**`CookiesList`

除列表格式，还能以其它格式返回：

* `cookies().as_str`：以`str`格式返回，只包含`name`和`value`字段，`'name1=value1; name2=value2'`格式
* `cookies().as_dict`：以`dict`格式返回，只包含`name`和`value`字段，`{'name1': 'value1', 'name2': 'value1'}`格式
* `cookies().as_json`：把列表转换为 json 返回

---

### 📌 `process_id`[​](#-process_id "-process_id的直接链接")

此属性返回浏览器进程 pid。

**类型：**`int`

---

### 📌 `states.is_alive`[​](#-statesis_alive "-statesis_alive的直接链接")

此属性返回浏览器是否仍可用。

**类型：**`bool`

---

### 📌 `states.is_existed`[​](#-statesis_existed "-statesis_existed的直接链接")

此属性返回浏览器是否接管的，而非本程序创建的。

**类型：**`bool`

---

### 📌 `states.is_headless`[​](#-statesis_headless "-statesis_headless的直接链接")

此属性返回浏览器是否无头模式。

**类型：**`bool`

---

### 📌 `states.is_incognito`[​](#-statesis_incognito "-statesis_incognito的直接链接")

此属性返回浏览器是否无痕模式。

**类型：**`bool`

---

✅️️ 其它浏览器行为[​](#️️-其它浏览器行为 "✅️️ 其它浏览器行为的直接链接")
----------------------------------------------

### 📌 `wait()`[​](#-wait "-wait的直接链接")

此方法用于等待若干秒。  
`scope`为`None`时，效果与`time.sleep()`没有区别，等待指定秒数。  
`scope`不为`None`时，获取两个参数之间的一个随机值，等待这个数值的秒数。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `second` | `float` | 必填 | 要等待的秒数，`scope`不为`None`时表示随机数范围起始值 |
| `scope` | `float` | `None` | 随机数范围结束值 |

| 返回类型 | 说明 |
| --- | --- |
| `Chromium` | 浏览器对象自身 |

---

### 📌 `wait.new_tab()`[​](#-waitnew_tab "-waitnew_tab的直接链接")

此方法用于等待新标签页出现。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` | `None` | 超时时间（秒），为`None`则使用对象`timeout`属性 |
| `curr_tab` | `str` `ChromiumTab` `MixTab` | `None` | 指定当前最新的 Tab 对象或标签页 id，用于判断新标签页出现，为`None`自动获取 |
| `raise_err` | `bool` | `None` | 等待失败时是否报错，为`None`时根据`Settings`设置 |

| 返回类型 | 说明 |
| --- | --- |
| `str` | 等待成功返回新标签页 id |
| `False` | 等待失败返回`False` |

---

### 📌 `wait.download_begin()`[​](#-waitdownload_begin "-waitdownload_begin的直接链接")

此方法用于等待浏览器下载开始。

有些下载是从临时闪现的标签页触发的，这种需要由浏览器对象去捕捉。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` | `None` | 超时时间（秒），`None`使用页面对象超时时间 |
| `cancel_it` | `bool` | `False` | 是否取消该任务 |

| 返回类型 | 说明 |
| --- | --- |
| `DownloadMission` | 等待成功返回下载任务对象 |
| `False` | 等待失败返回`False` |

---

### 📌 `wait.downloads_done()`[​](#-waitdownloads_done "-waitdownloads_done的直接链接")

此方法用于等待所有浏览器下载任务结束。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` | `None` | 超时时间（秒），为`None`时无限等待 |
| `cancel_if_timeout` | `bool` | `True` | 超时时是否取消剩余任务 |

| 返回类型 | 说明 |
| --- | --- |
| `bool` | 是否等待成功 |

---

### 📌 `clear_cache()`[​](#-clear_cache "-clear_cache的直接链接")

此方法用于清除缓存。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `cache` | `bool` | `True` | 是否清除缓存 |
| `cookies` | `bool` | `True` | 是否清除 cookies |

**返回：**`None`

---

### 📌 `reconnect()`[​](#-reconnect "-reconnect的直接链接")

此方法用于关闭与浏览器连接，并重新创建连接。

**参数：** 无

**返回：**`None`

---

### 📌 `quit()`[​](#-quit "-quit的直接链接")

此方法用于关闭浏览器。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` | `5` | 等待浏览器关闭超时时间（秒） |
| `force` | `bool` | `False` | 是否立刻强制终止进程 |
| `del_data` | `bool` | `False` | 是否删除用户文件夹 |

**返回：**`None`

[上一页

🛰️ 浏览器启动设置](/browser_control/browser_options)[下一页

🛰️ 标签页管理](/browser_control/tabs)

* [✅️️ 获取标签页对象或信息](#️️-获取标签页对象或信息)
  + [📌 `get_tab()`](#-get_tab)
  + [📌 `get_tabs()`](#-get_tabs)
  + [📌 `latest_tab`](#-latest_tab)
  + [📌 `tabs_count`](#-tabs_count)
  + [📌 `tab_ids`](#-tab_ids)
* [✅️️ 标签页操作](#️️-标签页操作)
  + [📌 `new_tab()`](#-new_tab)
  + [📌 `activate_tab()`](#-activate_tab)
  + [📌 `close_tabs()`](#-close_tabs)
  + [📌 单例模式说明](#-单例模式说明)
* [✅️️ 浏览器运行参数](#️️-浏览器运行参数)
  + [📌 `user_data_path`](#-user_data_path)
  + [📌 `download_path`](#-download_path)
  + [📌 几种超时参数](#-几种超时参数)
  + [📌 `timeout`](#-timeout)
  + [📌 `load_mode`](#-load_mode)
* [✅️️ 浏览器运行设置](#️️-浏览器运行设置)
  + [📌 `set.timeouts()`](#-settimeouts)
  + [📌 加载模式设置](#-加载模式设置)
  + [📌 `set.retry_times()`](#-setretry_times)
  + [📌 `set.retry_interval()`](#-setretry_interval)
  + [📌 `set.cookies()`](#-setcookies)
  + [📌 `set.cookies.clear()`](#-setcookiesclear)
  + [📌 `set.auto_handle_alert()`](#-setauto_handle_alert)
  + [📌 `set.download_path()`](#-setdownload_path)
  + [📌 `set.download_file_name()`](#-setdownload_file_name)
  + [📌 `set.when_download_file_exists()`](#-setwhen_download_file_exists)
  + [📌 `set.NoneElement_value()`](#-setnoneelement_value)
* [✅️️ 浏览器信息](#️️-浏览器信息)
  + [📌 `cookies()`](#-cookies)
  + [📌 `process_id`](#-process_id)
  + [📌 `states.is_alive`](#-statesis_alive)
  + [📌 `states.is_existed`](#-statesis_existed)
  + [📌 `states.is_headless`](#-statesis_headless)
  + [📌 `states.is_incognito`](#-statesis_incognito)
* [✅️️ 其它浏览器行为](#️️-其它浏览器行为)
  + [📌 `wait()`](#-wait)
  + [📌 `wait.new_tab()`](#-waitnew_tab)
  + [📌 `wait.download_begin()`](#-waitdownload_begin)
  + [📌 `wait.downloads_done()`](#-waitdownloads_done)
  + [📌 `clear_cache()`](#-clear_cache)
  + [📌 `reconnect()`](#-reconnect)
  + [📌 `quit()`](#-quit)

* 🚀 控制浏览器
* 🛰️ 浏览器启动设置

本页总览

🛰️ 浏览器启动设置
==========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

浏览器的启动配置非常繁杂，本库使用`ChromiumOptions`类管理启动配置，并且内置了常用配置的设置接口。

注意

该对象只能用于浏览器的启动，浏览器启动后，再修改该配置没有任何效果。接管已打开的浏览器时，启动配置也是无效的。

✅️️ 创建对象[​](#️️-创建对象 "✅️️ 创建对象的直接链接")
-------------------------------------

### 📌 导入[​](#-导入 "📌 导入的直接链接")

```
from DrissionPage import ChromiumOptions
```

---

### 📌 初始化参数[​](#-初始化参数 "📌 初始化参数的直接链接")

`ChromiumOptions`对象用于管理浏览器初始化配置。可从配置文件中读取配置来进行初始化。

| 初始化参数 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `read_file` | `bool` | `True` | 是否从 ini 文件中读取配置信息，为`False`则用默认配置创建 |
| `ini_path` | `Path` `str` | `None` | 指定 ini 文件路径，为`None`则读取内置 ini 文件 |

创建配置对象：

```
from DrissionPage import ChromiumOptions  
  
co = ChromiumOptions()
```

默认情况下，`ChromiumOptions`对象会从 ini 文件中读取配置信息，当指定`read_file`参数为`False`时，则以默认配置创建。

提醒

对象创建时已带有默认设置，如要清除，可调用`clear_arguments()`、`clear_prefs()`等方法。

---

✅️️ 使用方法[​](#️️-使用方法 "✅️️ 使用方法的直接链接")
-------------------------------------

创建配置对象后，可调整配置内容，然后在页面对象创建时以参数形式把配置对象传递进去，页面对象会根据配置对象的内容对浏览器进行初始化。

配置对象支持链式操作。

```
from DrissionPage import Chromium, ChromiumOptions  
  
# 创建配置对象（默认从 ini 文件中读取配置）  
co = ChromiumOptions()  
# 设置不加载图片、静音  
co.no_imgs(True).mute(True)  
  
# 以该配置创建页面对象  
page = Chromium(addr_or_opts=co)
```

```
from DrissionPage import  Chromium, ChromiumOptions  
  
co = ChromiumOptions()  
co.incognito()  # 匿名模式  
co.headless()  # 无头模式  
co.set_argument('--no-sandbox')  # 无沙盒模式  
page = Chromium(co)
```

---

✅️️ 命令行参数设置[​](#️️-命令行参数设置 "✅️️ 命令行参数设置的直接链接")
----------------------------------------------

Chromium 内核浏览器有一系列的启动配置，以`--`开头，可在浏览器创建时传入，控制浏览器行为和初始状态。

启动参数非常多，详见：[List of Chromium Command Line Switches](https://peter.sh/experiments/chromium-command-line-switches/)

### 📌 `set_argument()`[​](#-set_argument "-set_argument的直接链接")

此方法用于设置启动参数。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `arg` | `str` | 必填 | 启动参数名称 |
| `value` | `str` `None` `False` | `None` | 参数的值。带值的参数传入属性值，没有值的传入`None`。 如传入`False`  ，删除该参数。 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

**示例：** 无值和有值的参数设置

```
# 设置启动时最大化  
co.set_argument('--start-maximized')  
# 设置初始窗口大小  
co.set_argument('--window-size', '800,600')  
# 使用来宾模式打开浏览器  
co.set_argument('--guest')
```

---

### 📌 `remove_argument()`[​](#-remove_argument "-remove_argument的直接链接")

此方法用于在启动配置中删除一个启动参数，只要传入参数名称即可，不需要传入值。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `arg` | `str` | 必填 | 参数名称，有值的设置项传入设置名称即可 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象自身 |

**示例：** 删除无值和有值的参数设置

```
# 删除--start-maximized参数  
co.remove_argument('--start-maximized')  
# 删除--window-size参数  
co.remove_argument('--window-size')
```

---

### 📌 `clear_arguments()`[​](#-clear_arguments "-clear_arguments的直接链接")

此方法用于清空已设置的`arguments`参数。

**参数：** 无

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象自身 |

---

✅️️ 运行路径及端口[​](#️️-运行路径及端口 "✅️️ 运行路径及端口的直接链接")
----------------------------------------------

这部分是浏览器路径、用户文件夹路径和端口的设置。

### 📌 `set_browser_path()`[​](#-set_browser_path "-set_browser_path的直接链接")

此方法用于设置浏览器可执行文件路径。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `path` | `str` `Path` | 必填 | 浏览器文件路径 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

如果传入的字符串不是浏览器可执行文件路径，则会转为使用默认路径。

---

### 📌 `set_tmp_path()`[​](#-set_tmp_path "-set_tmp_path的直接链接")

此方法用于设置临时文件默认路径。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `path` | `str` `Path` | 必填 | 用户数据文件夹默认路径 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

---

### 📌 `set_local_port()`[​](#-set_local_port "-set_local_port的直接链接")

此方法用于设置本地启动端口。

与`set_address()`、`auto_port()`互斥。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `port` | `str` `int` | 必填 | 端口号 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

---

### 📌 `set_address()`[​](#-set_address "-set_address的直接链接")

此方法用于设置浏览器地址，支持 'ip:port' 格式和 ws 连接。

和`set_local_port()`、`auto_port()`互斥。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `address` | `str` | 必填 | 浏览器地址 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

---

### 📌 `auto_port()`[​](#-auto_port "-auto_port的直接链接")

此方法用于设置是否使用自动分配的端口，启动一个全新的浏览器。

如果设置为`True`，程序会自动寻找一个可用端口，并在指定路径或系统临时文件夹创建一个文件夹，用于储  存浏览器数据。

由于端口和用户文件夹都是唯一的，所以用这种方式启动的浏览器不会产生冲突，但也无法多次启动程序时重复接管同一个浏览器。

`set_local_port()`、`set_address()`和`set_user_data_path()`方法，和`auto_port()`互斥，即以后调用的为准。

注意

`auto_port()`支持多线程，但不支持多进程。  
多进程使用时，可用`scope`参数指定每个进程使用的端口范围，以免发生冲突。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_off` | `bool` | `True` | 是否开启自动分配端口和用户文件夹 |
| `scope` | `Tuple[int, int]` | `None` | 指定端口范围，不含最后的数字，为`None`则使用`[9600-19600)` |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

**示例：**

```
co.auto_port(True)
```

注意

启用此功能后即会获取端口和新建临时用户数据文件夹，若此时用`save()`方法保存配置到 ini 文件，ini 文件中的设置会被该端口和文件夹路径覆盖。这个覆盖对使用并没有很大影响。

---

### 📌 `set_user_data_path()`[​](#-set_user_data_path "-set_user_data_path的直接链接")

此方法用于设置用户文件夹路径。用户文件夹用于存储当前登陆浏览器的账号在使用浏览器时留下的痕迹，包括设置选项等。

一般来说用户文件夹的名称是 `User Data`。对于默认情况下的 Windows 中的 Chrome 浏览器来说，此文件夹位于 `%USERPROFILE%\AppData\Local\Google\Chrome\User Data\`，也就是当前系统登陆的用户目录的 `AppData` 内。实际情况可能有变，实际路径请在浏览器输入 `chrome://version/`，查阅其中的`个人资料路径`或者叫`用户配置路径`。若要使用独立的用户信息，可以将 `User Data` 目录整个复制到自定的其他位置，然后在代码中使用 `set_user_data_path()` 方法，参数填入自定义位置路径，这样便可使用独立的用户文件夹信息。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `path` | `str` `Path` | 必填 | 用户文件夹路径 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

---

### 📌 `use_system_user_path()`[​](#-use_system_user_path "-use_system_user_path的直接链接")

此方法设置是否使用系统安装的浏览器默认用户文件夹

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_off` | `bool` | `True` | `bool`表示开关 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

---

### 📌 `set_cache_path()`[​](#-set_cache_path "-set_cache_path的直接链接")

此方法用于设置缓存路径。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `path` | `str` `Path` | 必填 | 缓存路径 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

---

### 📌 `existing_only()`[​](#-existing_only "-existing_only的直接链接")

此方法设置是否  仅使用已启动的浏览器，如连接目标浏览器失败，会抛出异常，不会启动新浏览器。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_off` | `bool` | `True` | `bool`表示开关 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

---

  

✅️️ 使用插件[​](#️️-使用插件 "✅️️ 使用插件的直接链接")
-------------------------------------

`add_extension()`和`remove_extensions()`用于设置浏览器启动时要加载的插件。可以指定数量不限的插件。

### 📌 `add_extension()`[​](#-add_extension "-add_extension的直接链接")

此方法用于添加一个插件到浏览器。

插件是临时方式加载，不会保留在用户文件夹。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `path` | `str` `Path` | 必填 | 插件路径 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

Tips

根据作者的经验，把插件文件解压到一个独立文件夹，然后把插件路径指向这个文件夹，会比较稳定。

**示例：**

```
co.add_extension(r'D:\SwitchyOmega')
```

---

### 📌 `remove_extensions()`[​](#-remove_extensions "-remove_extensions的直接链接")

此方法用于移除配置对象中保存的所有插件路径。如需移除部分插件，请移除全部后再重新添加需要的插件。

**参数：** 无

**返回：** 配置对象自身

```
co.remove_extensions()
```

---

✅️️ 用户文件设置[​](#️️-用户文件设置 "✅️️ 用户文件设置的直接链接")
-------------------------------------------

除了启动参数，还有大量配置信息保存在浏览器的 `preferences` 文件中。

注意

`preferences` 文件是Chromium内核浏览器的配置信息文件，与 DrissionPage 的 `configs.ini` 完全不同。

以下方法用于对浏览器用户文件进行设置。

### 📌 `set_user()`[​](#-set_user "-set_user的直接链接")

Chromium 浏览器支持多用户配置，我们可以选择使用哪一个。默认为`'Default'`。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `user` | `str` | `'Default'` | 用户配置 文件夹名称 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

**示例：**

```
co.set_user(user='Profile 1')
```

---

### 📌 `set_pref()`[​](#-set_pref "-set_pref的直接链接")

此方法用于设置用户配置文件里的一个配置项。

在哪里可以查到所有的配置项？作者也没找到，知道的请告知。谢谢。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `arg` | `str` | 必填 | 设置项名称 |
| `value` | `str` | 必填 | 设置项值 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

**示例：**

```
# 禁止所有弹出窗口  
co.set_pref(arg='profile.default_content_settings.popups', value='0')  
# 隐藏是否保存密码的提示  
co.set_pref('credentials_enable_service', False)
```

---

### 📌 `remove_pref()`[​](#-remove_pref "-remove_pref的直接链接")

此方法用于在当前配置对象中删除一个`pref`配置项。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `arg` | `str` | 必填 | 设置项名称 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

**示例：**

```
co.remove_pref(arg='profile.default_content_settings.popups')
```

---

### 📌 `remove_pref_from_file()`[​](#-remove_pref_from_file "-remove_pref_from_file的直接链接")

此方法用于在用户配置文件删除一个配置项。注意与上一个方法不一样，如果用户配置文件中已经存在某个项，用`remove_pref()`
是不能删除的，只能用`remove_pref_from_file()`删除。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `arg` | `str` | 必填 | 设置项名称 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

**示例：**

```
co.remove_pref_from_file(arg='profile.default_content_settings.popups')
```

---

### 📌 `clear_prefs()`[​](#-clear_prefs "-clear_prefs的直接链接")

此方法用于清空已设置的`prefs`参数。

**参数：** 无

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象自身 |

---

✅️️ 运行参数设置[​](#️️-运行参数设置 "✅️️ 运行参数设置的直接链接")
-------------------------------------------

页面对象运行时需要用到的参数，也可以在`ChromiumOptions`中设置。

### 📌 `set_timeouts()`[​](#-set_timeouts "-set_timeouts的直接链接")

此方法用于设置几种超时时间，单位为秒。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `base` | `float` | `None` | 默认超时时间，用于元素等待、alert   等待、`WebPage`的 s 模式连接等等，除以下两个参数的场景，都使用这个设置 |
| `page_load` | `float` | `None` | 页面加载超时时间 |
| `script` | `float` | `None` | JavaScript 运行超时时间 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

**示例：**

```
co.set_timeouts(base=10)
```

---

### 📌 `set_retry()`[​](#-set_retry "-set_retry的直接链接")

此方法用于设置页面连接超时时的重试次数和间隔。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `times` | `int` | `None` | 连接失败重试次数 |
| `interval` | `float` | `None` | 连接失败重试间隔（秒） |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

---

### 📌 `set_load_mode()`[​](#-set_load_mode "-set_load_mode的直接链接")

此方法用于设置网页加载策略。

加载策略是指强制页面停止加载的时机，如加载完 DOM 即停止，不加载图片资源等，以提高自动化效率。

无论设置哪种策略，加载时间都不会超过`set_timeouts()`中`page_load`参数设置的时间。

加载策略：

* `'normal'`：阻塞进程，等待所有资源下载完成（默认）
* `'eager'`：DOM 就绪即停止加载
* `'none'`：网页连接成功即停止加载

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `value` | `str` | 必填 | 可接收`'normal'`、`'eager'`、`'none'` |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

**示例：**

```
co.set_load_mode('eager')
```

---

### 📌 `set_proxy()`[​](#-set_proxy "-set_proxy的直接链接")

该方法用于设置浏览器代理。

该设置在浏览器启动时一次性设置，设置后不能修改。且不支持带账号的代理。

如果需要运行时修改代理，或使用带账号的代理，可以用插件自行实现。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `proxy` | `str` | 必填 | 格式：协议://ip:port 当不指定协议时，默认使用 http 代理 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

**示例：**

```
co.set_proxy('http://localhost:1080')
```

---

### 📌 `set_download_path()`[​](#-set_download_path "-set_download_path的直接链接")

此方法用于设置下载文件保存路径。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `path` | `str` `Path` | 必填 | 下载路径 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

---

✅️️ 其它设置[​](#️️-其它设置 "✅️️ 其它设置的直接链接")
-------------------------------------

作者将一些常用的配置封装成方法，可以直接调用。

### 📌 `headless()`[​](#-headless "-headless的直接链接")

该方法用于设置是否以无界面模式启动浏览器。

如果指定端口已存在运行中的非无头浏  览器，会先关闭已有浏览器再启动新的。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_off` | `bool` | `True` | `True`和`False`表示开或关 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

**示例：**

```
co.headless(True)
```

---

### 📌 `new_env()`[​](#-new_env "-new_env的直接链接")

该方法用于设置是否使用全新环境创建浏览器。

如果指定端口已存在运行中的浏览器，会先关闭已有浏览器再启动新的。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_off` | `bool` | `True` | `True`和`False`表示开或关 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

---

### 📌 `set_flag()`[​](#-set_flag "-set_flag的直接链接")

此方法用于设置实验项，即`'chrome://flags'`中的项目。

设置无值的项，无须设置`value`参数，否则在该参数传入要设置的值。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `flag` | `str` | 必填 | 设置项名称 |
| `value` | `str` | `None` | 设置项值 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

**示例：**

```
from DrissionPage import ChromiumOptions  
  
co = ChromiumOptions()  
co.set_flag('temporary-unexpire-flags-m118', '1')  # 有值  
co.set_flag('disable-accelerated-2d-canvas')  # 无值
```

---

### 📌 `clear_flags_in_file()`[​](#-clear_flags_in_file "-clear_flags_in_file的直接链接")

此方法用于删除浏览器配置文件中已设置的实验项。

**参数：** 无

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

---

### 📌 `clear_flags()`[​](#-clear_flags "-clear_flags的直接链接")

此方法用于清空本对象中已设置的`flags`参数。

**参数：** 无

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象自身 |

---

### 📌 `incognito()`[​](#-incognito "-incognito的直接链接")

该方法用于设置是否以无痕模式启动浏览器。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_off` | `bool` | `True` | `True`和`False`表示开或关 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

---

### 📌 `ignore_certificate_errors()`[​](#-ignore_certificate_errors "-ignore_certificate_errors的直接链接")

该方法用于设置是否忽略证书错误。可以解决访问网页时出现的“您的连接不是私密连接”、“你的连接不是专用连接”等问题。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_off` | `bool` | `True` | `True`和`False`表示开或关 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

---

### 📌 `no_imgs()`[​](#-no_imgs "-no_imgs的直接链接")

该方法用于设置是否禁止加载图片。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_off` | `bool` | `True` | `True`和`False`表示开或关 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

**示例：**

```
co.no_imgs(True)
```

---

### 📌 `no_js()`[​](#-no_js "-no_js的直接链接")

该方法用于设置是否禁用 JavaScript。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_off` | `bool` | `True` | `True`和`False`表示开或关 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

**示例：**

```
co.no_js(True)
```

---

### 📌 `mute()`[​](#-mute "-mute的直接链接")

该方法用于设置是否静音。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_off` | `bool` | `True` | `True`和`False`表示开或关 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

**示例：**

```
co.mute(True)
```

---

### 📌 `set_user_agent()`[​](#-set_user_agent "-set_user_agent的直接链接")

该方法用于设置 user agent。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `user_agent` | `str` | 必填 | user agent文本 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumOptions` | 配置对象本身 |

**示例：**

```
co.set_user_agent(user_agent='Mozilla/5.0 (Macintos.....')
```

---

✅️️ 保存设置到文件[​](#️️-保存设置到文件 "✅️️ 保存设置到文件的直接链接")
----------------------------------------------

ini 文件是 DrissionPage 的配置文件，持久化记载一些配置参数。您可以把不同的配置保存到各自的 ini 文件，以便适应不同的场景。

### 📌 `save()`[​](#-save "-save的直接链接")

此方法用于保存配置项到一个 ini 文件。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `path` | `str` `Path` | `None` | ini 文件的路径， 传入`None`保存到当前读取的配置文件 |

| 返回类型 | 说明 |
| --- | --- |
| `str` | 保存的 ini 文件绝对路径 |

**示例：**

```
# 保存当前读取的ini文件  
co.save()  
  
# 把当前配置保存到指定的路径  
co.save(path=r'D:\tmp\settings.ini')
```

---

### 📌 `save_to_default()`[​](#-save_to_default "-save_to_default的直接链接")

此方法用于保存配置项到固定的默认 ini 文件。默认 ini 文件是指随 DrissionPage 内置的那个。

默认 ini 文件默认的路径是 Python 安装目录中的 `Lib\site-packages\DrissionPage\_configs\configs.ini`。

ini 文件初始内容点[这里](http://DrissionPage.cn/advance/ini)。

**参数：** 无

| 返回类型 | 说明 |
| --- | --- |
| `str` | 保存的 ini 文件绝对路径 |

**示例：**

```
co.save_to_default()
```

---

✅️️ `ChromiumOptions`属性[​](#️️-chromiumoptions属性 "️️-chromiumoptions属性的直接链接")
-----------------------------------------------------------------------------

### 📌 `address`[​](#-address "-address的直接链接")

该属性为要控制的浏览器地址，格式为 ip:port，默认为`'127.0.0.1:9222'`。

**类型：**`str`

---

### 📌 `browser_path`[​](#-browser_path "-browser_path的直接链接")

该属性返回浏览器可执行文件的路径。

**类型：**`str`

---

### 📌 `user_data_path`[​](#-user_data_path "-user_data_path的直接链接")

该 属性返回用户数据文件夹路径。

**类型：**`str`

---

### 📌 `tmp_path`[​](#-tmp_path "-tmp_path的直接链接")

该属性返回临时文件夹路径，可用于保存自动分配的用户文件夹路径。

**类型：**`str`

---

### 📌 `download_path`[​](#-download_path "-download_path的直接链接")

该属性返回默认下载路径文件路径。

**类型：**`str`

---

### 📌 `user`[​](#-user "-user的直接链接")

该属性返回用户配置文件夹名称。

**类型：**`str`

---

### 📌 `load_mode`[​](#-load_mode "-load_mode的直接链接")

该属性返回页面加载策略。有`'normal'`、`'eager'`、`'none'`三种

**类型：**`str`

---

### 📌 `timeouts`[​](#-timeouts "-timeouts的直接链接")

该属性返回超时设置。包括三种：`'base'`、`'page_load'`、`'script'`。

**类型：**`dict`

```
print(co.timeouts)
```

**输出：**

```
{  
    'base': 10,  
    'page_load': 30,  
    'script': 30  
}
```

---

### 📌 `retry_times`[​](#-retry_times "-retry_times的直接链接")

该属性返回连接失败时的重试次数。

**类型：**`int`

---

### 📌 `retry_interval`[​](#-retry_interval "-retry_interval的直接链接")

该属性返回连接失败时的重试间隔（秒）。

**类型：**`float`

---

### 📌 `proxy`[​](#-proxy "-proxy的直接链接")

该属性返回代理设置。

**类型：**`str`

---

### 📌 `arguments`[​](#-arguments "-arguments的直接链接")

该属性以`list`形式返回浏览器启动参数。

**类型：**`list`

---

### 📌 `extensions`[​](#-extensions "-extensions的直接链接")

该属性以`list`形式返回要加载的插件路径。

**类型：**`list`

---

### 📌 `preferences`[​](#-preferences "-preferences的直接链接")

该属性返回用户首选项配置。

**类型：**`dict`

---

### 📌 `system_user_path`[​](#-system_user_path "-system_user_path的直接链接")

该属性返回是否使用系统按照的浏览器的用户文件夹。

**类型：**`bool`

---

### 📌 `is_existing_only`[​](#-is_existing_only "-is_existing_only的直接链接")

该属性返回是否仅使用已打开的浏览器。

**类型：**`bool`

---

### 📌 `is_auto_port`[​](#-is_auto_port "-is_auto_port的直接链接")

该属性返回是否仅使用自动分配端口和用户文件夹路径。

**类型：**`bool`

---

### 📌 `is_headless`[​](#-is_headless "-is_headless的直接链接")

该属性返回是否以无头模式启动浏览器。

**类型：**`bool`

[上一页

🛰️ 连接浏览器](/browser_control/connect_browser)[下一页

🛰️ 浏览器对象](/browser_control/browser_object)

* [✅️️ 创建对象](#️️-创建对象)
  + [📌 导入](#-导入)
  + [📌 初始化参数](#-初始化参数)
* [✅️️ 使用方法](#️️-使用方法)
* [✅️️ 命令行参数设置](#️️-命令行参数设置)
  + [📌 `set_argument()`](#-set_argument)
  + [📌 `remove_argument()`](#-remove_argument)
  + [📌 `clear_arguments()`](#-clear_arguments)
* [✅️️ 运行路径及端口](#️️-运行路径及端口)
  + [📌 `set_browser_path()`](#-set_browser_path)
  + [📌 `set_tmp_path()`](#-set_tmp_path)
  + [📌 `set_local_port()`](#-set_local_port)
  + [📌 `set_address()`](#-set_address)
  + [📌 `auto_port()`](#-auto_port)
  + [📌 `set_user_data_path()`](#-set_user_data_path)
  + [📌 `use_system_user_path()`](#-use_system_user_path)
  + [📌 `set_cache_path()`](#-set_cache_path)
  + [📌 `existing_only()`](#-existing_only)
* [✅️️ 使用插件](#️️-使用插件)
  + [📌 `add_extension()`](#-add_extension)
  + [📌 `remove_extensions()`](#-remove_extensions)
* [✅️️ 用户文件设置](#️️-用户文件设置)
  + [📌 `set_user()`](#-set_user)
  + [📌 `set_pref()`](#-set_pref)
  + [📌 `remove_pref()`](#-remove_pref)
  + [📌 `remove_pref_from_file()`](#-remove_pref_from_file)
  + [📌 `clear_prefs()`](#-clear_prefs)
* [✅️️ 运行参数设置](#️️-运行参数设置)
  + [📌 `set_timeouts()`](#-set_timeouts)
  + [📌 `set_retry()`](#-set_retry)
  + [📌 `set_load_mode()`](#-set_load_mode)
  + [📌 `set_proxy()`](#-set_proxy)
  + [📌 `set_download_path()`](#-set_download_path)
* [✅️️ 其它设置](#️️-其它设置)
  + [📌 `headless()`](#-headless)
  + [📌 `new_env()`](#-new_env)
  + [📌 `set_flag()`](#-set_flag)
  + [📌 `clear_flags_in_file()`](#-clear_flags_in_file)
  + [📌 `clear_flags()`](#-clear_flags)
  + [📌 `incognito()`](#-incognito)
  + [📌 `ignore_certificate_errors()`](#-ignore_certificate_errors)
  + [📌 `no_imgs()`](#-no_imgs)
  + [📌 `no_js()`](#-no_js)
  + [📌 `mute()`](#-mute)
  + [📌 `set_user_agent()`](#-set_user_agent)
* [✅️️ 保存设置到文件](#️️-保存设置到文件)
  + [📌 `save()`](#-save)
  + [📌 `save_to_default()`](#-save_to_default)
* [✅️️ `ChromiumOptions`属性](#️️-chromiumoptions属性)
  + [📌 `address`](#-address)
  + [📌 `browser_path`](#-browser_path)
  + [📌 `user_data_path`](#-user_data_path)
  + [📌 `tmp_path`](#-tmp_path)
  + [📌 `download_path`](#-download_path)
  + [📌 `user`](#-user)
  + [📌 `load_mode`](#-load_mode)
  + [📌 `timeouts`](#-timeouts)
  + [📌 `retry_times`](#-retry_times)
  + [📌 `retry_interval`](#-retry_interval)
  + [📌 `proxy`](#-proxy)
  + [📌 `arguments`](#-arguments)
  + [📌 `extensions`](#-extensions)
  + [📌 `preferences`](#-preferences)
  + [📌 `system_user_path`](#-system_user_path)
  + [📌 `is_existing_only`](#-is_existing_only)
  + [📌 `is_auto_port`](#-is_auto_port)
  + [📌 `is_headless`](#-is_headless)

* 🚀 控制浏览器
* 🛰️ 连接浏览器

本页总览

🛰️ 连接浏览器
========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

`Chromium`对象用于连接和管理浏览器。标签页的开关和获取、整体运行参数配置、浏览器信息获取等都由它进行。

根据不同的配置，可以接管已打开的浏览器，也可以启动新的浏览器。

每个浏览器只能有一个`Chromium`对象（同一进程中）。对同一个浏览器重复使用`Chromium()`获取的都是同一个对象。

Tips

程序结束时，被打开的浏览器不会主动关闭（VSCode 启动的除外），以便下次运行程序时使用。
新手在使用无头模式时需注意，程序关闭后其实浏览器进程还在，只是看不见。

✅️ `Chromium`初始化参数[​](#️-chromium初始化参数 "️-chromium初始化参数的直接链接")
--------------------------------------------------------------

| 初始化参数 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `addr_or_opts` | `str` `int` `ChromiumOptions` | `None` | 浏览器启动配置或接管信息。 传入 'ip: port' 字符串、浏览器 ws 地址、端口数字或`ChromiumOptions`对象时按配置启动或接管浏览器； 为`None`时使用配置文件配置启动浏览器 |
| `session_options` | `SessionOptions` `None` `False` | `None` | 使用双模 Tab 时使用的默认 Session 配置，为`None`使用 ini 文件配置，为`False`不从 ini 读取 |

---

✅️ 直接创建[​](#️-直接创建 "✅️ 直接创建的直接链接")
----------------------------------

### 📌 默认方式[​](#-默认方式 "📌 默认方式的直接链接")

这种方式代码最简洁，程序会使用默认配置，自动生成页面对象。

```
from DrissionPage import Chromium  
  
browser = Chromium()
```

创建`Chromium`对象时会在指定端口启动浏览器，或接管该端口已有浏览器。

默认情况下，程序使用 9222 端口，浏览器可执行文件路径为`'chrome'`。

如路径中没找到浏览器可执行文件，Windows 系统下程序会在注册表中查找路径。

如果都没找到，则要用下文介绍的手动配置方法。

直接创建时，程序默认读取 ini 文件配置，如 ini 文件不存在，会使用内置配置。

默认 ini 和内置配置信息详见“进阶使用->配置文件的使用”章节。

Tips

您可以修改配置文件中的配置，实现所有程序都按您的需要进行启动，详见”启动配置“章节。

---

### 📌 指定端口或地址[​](#-指定端口或地址 "📌 指定端口或地址的直接链接")

创建`Chromium`对象时向`addr_or_opts`参数传入端口号或地址，可接管指定端口浏览器，若端口空闲，使用默认配置在该端口启动一个浏览器。

传入端口时用`int`类型，传入地址时用`'ip:port'`格式，传入 ws 地址需为完整地址。

```
from DrissionPage import Chromium  
  
browser = Chromium(9333)  # 接管9333端口的浏览器，如该端口空闲，启动一个浏览器  
browser = Chromium('127.0.0.1:9333')  # 与上一行一样  
browser = Chromium('ws://127.0.0.1:8987/devtools/browser/3e590fc5-4587-47e1-8756-cf6784f2fef3')  # 指定ws连接
```

---

✅️ 通过配置信息创建[​](#️-通过配置信息创建 "✅️ 通过配置信息创建的直接链接")
----------------------------------------------

如果需要已指定方式启动浏览器，可使用`ChromiumOptions`。它是专门用于设置浏览器初始状态的类，内置了常用的配置。详细使用方法见“浏览器启动配置”一节。

### 📌 使用方法[​](#-使用方法 "📌 使用方法的直接链接")

`ChromiumOptions`用于管理创建浏览器时的配置，内置了常用的配置，并能实现链式操作。详细使用方法见“启动配置”一节。

| 初始化参数 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `read_file` | `bool` | `True` | 是否从 ini 文件中读取配置信息，如果为`False`则用默认配置创建 |
| `ini_path` | `str` | `None` | 文件路径，为`None`则读取默认 ini 文件 |

注意

* 配置对象只有在启动浏览器时生效。
* 浏览器创建后再修改这个配置是没有效果的。
* 接管已打开的浏览器配置也不会生效。

```
# 导入 ChromiumOptions  
from DrissionPage import Chromium, ChromiumOptions  
  
# 创建浏览器配置对象，指定浏览器路径  
co = ChromiumOptions().set_browser_path(r'D:\chrome.exe')  
# 用该配置创建页面对象  
browser = Chromium(addr_or_opts=co)
```

---

### 📌 使用指定 ini 文件创建[​](#-使用指定-ini-文件创建 "📌 使用指定 ini 文件创建的直接链接")

以上方法是使用默认 ini 文件中保存的配置信息创建对象，你可以保存一个 ini 文件到别的地方，并在创建对象时指定使用它。

```
from DrissionPage import Chromium, ChromiumOptions  
  
# 创建配置对象时指定要读取的ini文件路径  
co = ChromiumOptions(ini_path=r'./config1.ini')  
# 使用该配置对象创建页面  
browser = Chromium(addr_or_opts=co)
```

---

  

✅️ 接管已打开的浏览器[​](#️-接管已打开的浏览器 "✅️ 接管已打开的浏览器的直接链接")
-------------------------------------------------

页面对象创建时，只要指定的地址（'ip:port' 或 ws 地址）已有浏览器在运行，就会直接接管。无论浏览器是下面哪种方式启动的。

### 📌 用程序启动的浏览器[​](#-用程序启动的浏览器 "📌 用程序启动的浏览器的直接链接")

默认情况下，创建浏览器页面对象时会自动启动一个浏览器。只要这个浏览器不关闭，下次运行程序时会接管同一个浏览器继续操作（配置的 ip:port 信息不变）。

这种方式极大地方便了程序的调试，使程序不必每次重新开始，可以单独调试某个功能。

```
from DrissionPage import Chromium  
  
# 在9333端口启动浏览器同时创建对象，如果浏览器已经存在，则接管它  
browser = Chromium(9333)
```

---

### 📌 手动打开的浏览器[​](#-手动打开的浏览器 "📌 手动打开的浏览器的直接链接")

如果需要手动打开浏览器再接管，可以这样做：

1. 右键点击浏览器图标，选择属性
2. 在“目标”路径后面加上 `--remote-debugging-port=端口号`（注意最前面有个空格）
3. 点击确定
4. 在程序中的浏览器配置中指定接管该端口浏览器

文件快捷方式的目标路径设置：

```
"D:\chrome.exe" --remote-debugging-port=9333
```

程序代码：

```
from DrissionPage import Chromium  
  
browser = Chromium(9333)
```

注意

接管浏览器时只有`local_port`、`address`参数是有效的。

---

### 📌 bat 文件启动的浏览器[​](#-bat-文件启动的浏览器 "📌 bat 文件启动的浏览器的直接链接")

可以把上一种方式的目标路径设置写进 bat 文件（Windows系统），运行 bat 文件来启动浏览器，再用程序接管。

新建一个文本文件，在里面输入以下内容（路径改为自己电脑的）：

```
"D:\chrome.exe" --remote-debugging-port=9333
```

保存后把后缀改成 bat，然后双击运行就能在 9333 端口启动一个浏览器。程序代码则和上一个方法一致。

---

### 📌 用 ws 连接的远程浏览器[​](#-用-ws-连接的远程浏览器 "📌 用 ws 连接的远程浏览器的直接链接")

直接使用 ws 地址：

```
from DrissionPage import Chromium  
  
browser = Chromium('wss://****.com/ws?apiKey=5482a4cba773')
```

在`ChromiumOptions`设置 ws 地址：

```
from DrissionPage import Chromium, ChromiumOptions  
  
co = ChromiumOptions().set_address('wss://****.com/ws?apiKey=5482a4cba773')  
browser = Chromium(co)
```

---

✅️ 多浏览器共存[​](#️-多浏览器共存 "✅️ 多浏览器共存的直接链接")
----------------------------------------

如果想要同时操作多个浏览器，或者自己在使用其中一个上网，同时控制另外几个跑自动化，就需要给这些被程序控制的浏览器设置单独的 **端口** 和 **用户文件夹**，否则会造成冲突。

### 📌 指定独立端口和数据文件夹[​](#-指定独立端口和数据文件夹 "📌 指定独立端口和数据文件夹的直接链接")

每个要启动的浏览器使用一个独立的`ChromiumOptions`对象进行设置：

```
from DrissionPage import Chromium, ChromiumOptions  
  
# 创建多个配置对象，每个指定不同的端口号和用户文件夹路径  
co1 = ChromiumOptions().set_local_port(9111).set_user_data_path(r'D:\data1')  
co2 = ChromiumOptions().set_local_port(9222).set_user_data_path(r'D:\data2')  
  
# 创建多个页面对象  
tab1 = Chromium(addr_or_opts=co1).latest_tab  
tab2 = Chromium(addr_or_opts=co2).latest_tab  
  
# 每个页面对象控制一个浏览器  
tab1.get('http://DrissionPage.cn')  
tab2.get('https://www.baidu.com')
```

注意

每个浏览器都要设置独立的端口号和用户文件夹，二者缺一不可。

---

### 📌 `auto_port()`方法[​](#-auto_port方法 "-auto_port方法的直接链接")

`ChromiumOptions`对象的`auto_port()`方法，可以指定程序每次使用空闲的端口和临时用户文件夹创建浏览器。

使用`auto_port()`的配置对象可由多个`Chromium`对象共用，不会出现冲突。

这种方式创建的浏览器是全新不带任何数据的，并且运行数据会自动清除。

Tips

`auto_port()`支持多线程，多进程使用时由小概率出现端口冲突。  
多进程使用时，可用`scope`参数指定每个进程使用的端口范围，以免发生冲突。

```
from DrissionPage import Chromium, ChromiumOptions  
  
co = ChromiumOptions().auto_port()  
  
tab1 = Chromium(addr_or_opts=co).latest_tab  
tab2 = Chromium(addr_or_opts=co).latest_tab  
  
tab2.get('http://DrissionPage.cn')  
tab1.get('https://www.baidu.com')
```

---

### 📌 在 ini 文件设置自动分配[​](#-在-ini-文件设置自动分配 "📌 在 ini 文件设置自动分配的直接链接")

可以把自动分配的配置记录到 ini 文件，这样无需创建`ChromiumOptions`，每次启动的浏览器都是独立的，不会冲突。但和`auto_port()`一样，这些浏览器也不能复用。

```
from DrissionPage import ChromiumOptions  
  
ChromiumOptions().auto_port(True).save()
```

这段代码把该配置记录到 ini 文件，只需运行一次，要关闭的话把参数换成`False`再执行一次即可。

```
from DrissionPage import Chromium  
  
tab1 = Chromium().latest_tab  
tab2 = Chromium().latest_tab  
  
tab1.get('http://DrissionPage.cn')  
tab2.get('https://www.baidu.com')
```

---

✅️ 使用系统浏览器用户目录[​](#️-使用系统浏览器用户目录 "✅️ 使用系统浏览器用户目录的直接链接")
-------------------------------------------------------

初始默认配置下，程序会为每个使用的端口创建空的用户目录，并且每次接管都使用，这样可以有效避免浏览器冲突。

有些时候我们希望使用系统安装的浏览器的默认用户文件夹。以便复用用户信息和插件等。

我们可以这样设置：

### 📌 使用`ChromiumOptions`[​](#-使用chromiumoptions "-使用chromiumoptions的直接链接")

用`ChromiumOptions`在每次启动时配置。

注意

使用这种方法时，需关闭已启动的系统浏览器，否则会连接失败。

```
from DrissionPage import Chromium, ChromiumOptions  
  
co = ChromiumOptions().use_system_user_path()  
browser = Chromium(co)
```

---

### 📌 使用 ini 文件[​](#-使用-ini-文件 "📌 使用 ini 文件的直接链接")

把这个配置记录到 ini 文件，就不用每次使用都配置。

注意

使用这种方法时，需关闭已启动的系统浏览器，否则会连接失败。

```
from DrissionPage import ChromiumOptions  
  
ChromiumOptions().use_system_user_path().save()
```

---

### 📌 手动打开再接管[​](#-手动打开再接管 "📌 手动打开再接管的直接链接")

参考上文 “接管已打开浏览器” 的方法，手动为浏览器设置端口启动，再用 DrissionPage 接管。

```
from DrissionPage import Chromium  
  
browser = Chromium(9333)  # 已手动在9333端口启动浏览器
```

---

✅️ 创建全新的浏览器[​](#️-创建全新的浏览器 "✅️ 创建全新的浏览器的直接链接")
----------------------------------------------

默认情况下，程序会复用之前用过的浏览器用户数据，因此可能带有登录数据、历史记录等。

如果想打开全新的浏览器，可用以下方法：

### 📌 使用`auto_port()`[​](#-使用auto_port "-使用auto_port的直接链接")

上文提过的`auto_port()`方法，会自动查找一个空闲的端口启动全新的浏览器。

示例见上文。

---

### 📌 使用`new_env()`[​](#-使用new_env "-使用new_env的直接链接")

`ChromiumOptions`对象的`new_env()`方法，可指定启动全新的浏览器。

如果指定端口已有浏览器，会自动关闭再启动新的。

```
from DrissionPage import Chromium, ChromiumOptions  
  
co = ChromiumOptions().new_env()  
browser = Chromium(co)
```

---

### 📌 手动指定端口和路径[​](#-手动指定端口和路径 "📌 手动指定端口和路径的直接链接")

给浏览器用户文件夹路径指定空的路径，以及指定一个空闲的端口，即可打开全新浏览器。

```
from DrissionPage import Chromium, ChromiumOptions  
  
co = ChromiumOptions().set_local_port(9333).set_user_data_path(r'C:\tmp')  
browser = Chromium(co)
```

---

✅️ 用户文件夹位置[​](#️-用户文件夹位置 "✅️ 用户文件夹位置的直接链接")
-------------------------------------------

复用用户文件夹可使用已登录的状态、已安装的插件、已设置好的配置等。

以下不同配置下用户文件夹的存放位置。

### 📌 默认配置[​](#-默认配置 "📌 默认配置的直接链接")

默认配置下，由 DrissionPage 创建的浏览器，用户文件夹在系统临时文件夹的`DrissionPage\userData`文件夹内，以端口命名。

假如用 DrissionPage 默认配置在 9222 端口创建一个浏览器，那么用户数据就存放在`C:\Users\用户名\AppData\Local\Temp\DrissionPage\userData\9222`路径。

这个用户文件夹不会主动清除，下次再使用 9222 端口时，会继续使用。

如果使用`auto_port()`，会存放在系统临时文件夹的`DrissionPage\autoPortData`文件夹内，以端口命名。

如`C:\Users\用户名\AppData\Local\Temp\DrissionPage\autoPortData\21489`。

这个用户文件夹是临时的，用完会被主动清除。

---

### 📌 自定义位置[​](#-自定义位置 "📌 自定义位置的直接链接")

如果要指定用户文件夹存放位置，可用`ChromiumOptions`对象的`set_tmp_path()`方法。

也可以保持到 ini 文件，可省略每次设置。

示例：

```
from DrissionPage import ChromiumOptions  
  
ChromiumOptions().set_tmp_path(r'D:\tmp').save()  # 保存到ini文件
```

---

### 📌 单独指定某个用户文件夹[​](#-单独指定某个用户文件夹 "📌 单独指定某个用户文件夹的直接链接")

指定用户文件夹路径，或使用系统文件夹路径，请查看上文。

```
from DrissionPage import Chromium, ChromiumOptions  
  
co = ChromiumOptions().set_user_data_path(r'D:\tmp')  
browser = Chromium(co)
```

```
from DrissionPage import Chromium, ChromiumOptions  
  
co = ChromiumOptions().use_system_user_path()  
browser = Chromium(co)
```

[上一页

🛰️ 概述](/browser_control/intro)[下一页

🛰️ 浏览器启动设置](/browser_control/browser_options)

* [✅️ `Chromium`初始化参数](#️-chromium初始化参数)
* [✅️ 直接创建](#️-直接创建)
  + [📌 默认方式](#-默认方式)
  + [📌 指定 端口或地址](#-指定端口或地址)
* [✅️ 通过配置信息创建](#️-通过配置信息创建)
  + [📌 使用方法](#-使用方法)
  + [📌 使用指定 ini 文件创建](#-使用指定-ini-文件创建)
* [✅️ 接管已打开的浏览器](#️-接管已打开的浏览器)
  + [📌 用程序启动的浏览器](#-用程序启动的浏览器)
  + [📌 手动打开的浏览器](#-手动打开的浏览器)
  + [📌 bat 文件启动的浏览器](#-bat-文件启动的浏览器)
  + [📌 用 ws 连接的远程浏览器](#-用-ws-连接的远程浏览器)
* [✅️ 多浏览器共存](#️-多浏览器共存)
  + [📌 指定独立端口和数据文件夹](#-指定独立端口和数据文件夹)
  + [📌 `auto_port()`方法](#-auto_port方法)
  + [📌 在 ini 文件设置自动分配](#-在-ini-文件设置自动分配)
* [✅️ 使用系统浏览器用户目录](#️-使用系统浏览器用户目录)
  + [📌 使用`ChromiumOptions`](#-使用chromiumoptions)
  + [📌 使用 ini 文件](#-使用-ini-文件)
  + [📌 手动打开再接管](#-手动打开再接管)
* [✅️ 创建全新的浏览器](#️-创建全新的浏览器)
  + [📌 使用`auto_port()`](#-使用auto_port)
  + [📌 使用`new_env()`](#-使用new_env)
  + [📌 手动指定端口和路径](#-手动指定端口和路径)
* [✅️ 用户文件夹位置](#️-用户文件夹位置)
  + [📌 默认配置](#-默认配置)
  + [📌 自定义位置](#-自定义位置)
  + [📌 单独指定某个用户文件夹](#-单独指定某个用户文件夹)

* 🚀 控制浏览器
* 🛰️ 获取控制台信息

本页总览

🛰️ 获取控制台信息
==========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

获取控制台信息的逻辑和监听网络数据差不多，是通过监听控制台数据实现的。

注意

不是所有显示在控制台的信息都能获取，需要用`console.log()`等方法输出到控制台的才能获取。

✅️ 示例[​](#️-示例 "✅️ 示例的直接链接")
----------------------------

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.console.start()  
tab.run_js('console.log("DrissionPage");')  
data = tab.console.wait()  
print(data.text)  # 输出：DrissionPage
```

---

✅️ 启动和停止[​](#️-启动和停止 "✅️ 启动和停止的直接链接")
-------------------------------------

### 📌 `console.start()`[​](#-consolestart "-consolestart的直接链接")

此方法用于启动控制台信息监听。

**参数：** 无

**返回：**`None`

---

### 📌 `console.stop()`[​](#-consolestop "-consolestop  的直接链接")

此方法用于停止监听，清空已监听到的信息列表。

**参数：** 无

**返回：**`None`

---

  

✅️ 获取信息[​](#️-获取信息 "✅️ 获取信息的直接链接")
----------------------------------

### 📌 `console.wait()`[​](#-consolewait "-consolewait的直接链接")

此方法用于等待一条控制台信息。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` `None` | `None` | 超时时间（秒），为`None`无限等待 |

| 返回类型 | 说明 |
| --- | --- |
| `ConsoleData` | 控制台信息数据包对象 |
| `False` | 等待超时时 |

---

### 📌 `console.steps()`[​](#-consolesteps "-consolesteps的直接链接")

此方法返回一个可迭代对象，用于`for`循环，每次循环可从中获取到的信息。

可实现实时获取并返回数据包。

如果`timeout`超时，会中断循环。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` `None` | `None` | 每个信息等待时间（秒），为`None`表示无限等待 |

| 返回类型 | 说明 |
| --- | --- |
| `ConsoleData` | 控制台信息数据包对象 |

---

### 📌 `console.messages`[​](#-consolemessages "-consolemessages的直接链接")

此属性以`list`方式返回获取到的信息，返回后会清空列表。

| 返回类型 | 说明 |
| --- | --- |
| `List[ConsoleData]` | 控制台信息对象组成的列表 |

---

✅️ 其它[​](#️-其它 "✅️ 其它的直接链接")
----------------------------

### 📌 `console.listening`[​](#-consolelistening "-consolelistening的直接链接")

此属性返回监听是否进行中。

**返回：** `bool`

---

### 📌 `console.clear()`[​](#-consoleclear "-consoleclear的直接  链接")

此方法用于清空已获取但未返回的信息。

**参数：** 无

**返回：**`None`

---

✅️ `ConsoleData`对象[​](#️-consoledata对象 "️-consoledata对象的直接链接")
--------------------------------------------------------------

`ConsoleData`对象是获取到的数据包结果对象，包含了数据包各种信息。

### 📌 `对象属性`[​](#-对象属性 "-对象属性的直接链接")

| 属性名称 | 数据类型 | 说明 |
| --- | --- | --- |
| `source` | `str` | 来源 |
| `level` | `str` | 类型 |
| `text` | `str` | 内容文本 |
| `body` | `Any` | 把`text`进行 json 解析 |
| `url` | `str` | 网址 |
| `line` | `str` | 行号 |
| `column` | `str` | 列号 |

[上一页

🛰️ 监听网络数据](/browser_control/listener)[下一页

🛰️ 截图和录像](/browser_control/screen)

* [✅️ 示例](#️-示例)
* [✅️ 启动和停止](#️-启动和停止)
  + [📌 `console.start()`](#-consolestart)
  + [📌 `console.stop()`](#-consolestop)
* [✅️ 获取信息](#️-获取信息)
  + [📌 `console.wait()`](#-consolewait)
  + [📌 `console.steps()`](#-consolesteps)
  + [📌 `console.messages`](#-consolemessages)
* [✅️ 其它](#️-其它)
  + [📌 `console.listening`](#-consolelistening)
  + [📌 `console.clear()`](#-consoleclear)
* [✅️ `ConsoleData`对象](#️-consoledata对象)
  + [📌 `对象属性`](#-对象属性)

* 🚀 控制浏览器
* 🛰️ 元素交互

本页总览

🛰️ 元素交互
=======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

本节介绍与浏览器元素的交互。浏览器元素对象为`ChromiumElement`。

✅️️ 点击元素[​](#️️-点击元素 "✅️️ 点击元素的直接链接")
-------------------------------------

### 📌 `click()`和`click.left()`[​](#-click和clickleft "-click和clickleft的直接链接")

这两个方法作用是一样的，用于左键点击元素。可选择模拟点击或 js 点击。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `by_js` | `bool` | `False` | 指定点击行为方式。 为`None`时，如不被遮挡，用模拟点击，否则用 js 点击 为`True`时直接用 js 点击； 为`False`时强制模拟点击，被遮挡也会进行点击 |
| `timeout` | `float` | `1.5` | 模拟点击的超时时间（秒），等待元素可见、可用、进入视口 |
| `wait_stop` | `bool` | `True` | 点击前是否等待元素停止运动 |

| 返回值 | 说明 |
| --- | --- |
| `False` | `by_js`为`False`，且元素不可用、不可见时，返回`False` |
| `True` | 除以上情况，其余情况都返回`True` |

**示例：**

```
# 对ele元素进行模拟点击，如判断被遮挡也会点击  
ele.click()  
  
# 用js方式点击ele元素，无视遮罩层  
ele.click(by_js=True)  
  
# 如元素不被遮挡，用模拟点击，否则用js点击  
ele.click(by_js=None)
```

默认情况下，`by_js`为`None`，优先用模拟方式点击，如遇遮挡、元素不可用、不可见、无法自动进入视口，等待直到超时后自动改用 js
方式点击。

`by_js`为`False`，程序会强制使用模拟点击，即使被遮挡也会点击元素位  置。如果元素不可见、不可用，会返回`False`。如元素无法自动滚动到视口，会改用
js 点击。

`by_js`为`True`时，则可无视任何遮挡，只要元素在 DOM 内，就能点击得到，但元素是否响应点击视网页所用架构而定。

可以根据需要灵活地对元素进行操作。

在模拟点击前，程序会先尝试把元素滚动到视口中。

默认情况下，如无法进行模拟点击（元素无法进入视口、不可用、隐藏）时，左键单击会返回`False`。但也可以通过全局设置使其抛出异常：

```
from DrissionPage.common import Settings  
  
Settings.set_raise_click_failed(True)  
ele.click()  # 如无法点击则抛出异常
```

---

### 📌 `click.right()`[​](#-clickright "-clickright的直接链接")

此方法实现右键单击元素。

**参数：** 无

**返回：**`None`

**示例：**

```
ele.click.right()
```

---

### 📌 `click.middle()`[​](#-clickmiddle "-clickmiddle的直接链接")

此方法实现中键单击元素。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `get_tab` | `bool` | `True` | 是否返回新出现的 Tab 对象 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumTab` | 如`get_tab`参数为`True`，元素在`ChromiumTab`返回对象 |
| `MixTab` | 如`get_tab`参数为`True`，元素在`MixTab`里时返回的对象 |
| `None` | `get_tab`参数为`False`时 |

**示例：**

```
tab = ele.click.middle()  
print(tab.title)
```

---

### 📌 `click.multi()`[​](#-clickmulti "-clickmulti的直接链接")

此方法实现左键多次点击元素。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `times` | `int` | `2` | 点击次数 |

**返回：**`None`

---

### 📌 `click.at()`[​](#-clickat "-clickat的直接链接")

此方法用于带偏移量点击元素，偏移量相对于元素左上角坐标。不传入`offset_x`和`offset_y`时点击元素中间点。  
点击的目标不一定在元素上，可以传入负值，或大于元素大小的值，点击元素附近的区域。向右和向下为正值，向左和向上为负值。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `offset_x` | `float` | `None` | 相对元素左上角坐标的 x 轴偏移量，向下向右为正 |
| `offset_y` | `float` | `None` | 相对元素左上角坐标的 y 轴偏移量，向下向右为  正 |
| `button` | `str` | `'left'` | 要点击的键，传入`'left'`、`'right'`、`'middle'`、`'back'`、`'forward'` |
| `count` | `int` | `1` | 点击次数 |

**返回：**`None`

**示例：**

```
# 点击元素右上方 50*50 的位置  
ele.click.at(50, -50)  
  
# 点击元素上中部，x相对左上角向右偏移50，y保持在元素中点  
ele.click.at(offset_x=50)  
  
# 和click()一致，但没有重试功能  
ele.click.at()
```

---

### 📌 `click.to_upload()`[​](#-clickto_upload "-clickto_upload的直接链接")

此方法用于点击元素，触发文件选择框并把指定的文件路径添加到网页，详见“文件上传”章节。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `file_paths` | `str` `Path` `list` `tuple` | 必填 | 文件路径，如果上传框支持多文件，可传入列表或字符串，字符串时多个文件用`\n`分隔 |
| `by_js` | `bool` | `False` | 是否用 js 方式点击，逻辑与`click()`一致 |

**返回：**`None`

---

### 📌 `click.to_download()`[​](#-clickto_download "-clickto_download的直接链接")

此方法用于点击元素触发下载，并返回下载任务对象。用法详见“文件下载”章节。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `save_path` | `str` `Path` | 必填 | 保存路径，为`None`保存在原来设置的，如未设置保存到当前路径 |
| `rename` | `str` | `None` | 重命名文件名，为`None`则不修改 |
| `suffix` | `str` | `'left'` | 指定文件后缀，为`None`则不修改 |
| `new_tab` | `bool` | `1` | 该下载是否在新 tab 中触发 |
| `by_js` | `bool` | `False` | 是否用 js 方式点击，逻辑与`click()`一致 |
| `timeout` | `float` | `None` | 超时时间（秒），为`None`时使用页面对象默认超时时间 |

| 返回类型 | 说明 |
| --- | --- |
| `DownloadMission` | 下载任务对象 |

---

### 📌 `click.for_new_tab()`[​](#-clickfor_new_tab "-clickfor_new_tab的直接链接")

在预期点击后会出现新 tab 的时候，  可用此方法点击，会等待并返回新 tab 对象。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `by_js` | `bool` | `False` | 是否用 js 方式点击，逻辑与`click()`一致 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumTab` | 元素在`ChromiumTab`里时返回 |
| `MixTab` | 元素在`MixTab`里时返回 |

---

✅️️ 输入内容[​](#️️-输入内容 "✅️️ 输入内容的直接链接")
-------------------------------------

### 📌 `clear()`[​](#-clear "-clear的直接链接")

此方法用于清空元素文本，可选择模拟按键或 js 方式。

模拟按键方式会自动输入`ctrl-a-del`组合键来清除文本框，js 方式则直接把元素`value`属性设置为`''`。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `by_js` | `bool` | `False` | 是否用 js 方式清空 |

**返回：**`None`

**示例：**

```
ele.clear()
```

---

### 📌 `input()`[​](#-input "-input的直接链接")

此方法用于向元素输入文本或组合键，也可用于输入文件路径到上传控件。可选择输入前是否清空元素。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `vals` | `Any` | `False` | 文本值或按键组合 对文件上传控件时输入路径字符串或其组成的列表 |
| `clear` | `bool` | `False` | 输入前是否清空文本框 |
| `by_js` | `bool` | `False` | 是否用 js 方式输入，为`True`时不能输入组合键 |

**返回：**`None`

Tips

* 有些文本框可以接收回车代替点击按钮，可以直接在文本末尾加上`'\n'`。
* 会自动把非`str`数据转换为`str`。

**示例：**

```
# 输入文本  
ele.input('Hello world!')  
  
# 输入文本并回车  
ele.input('Hello world!\n')
```

---

### 📌 输入组合键[​](#-输入组合键 "📌 输入组合键的直接链接")

使用组合键或要传入特殊按键前，先要导入按键类`Keys`。

```
from DrissionPage.common import Keys
```

然后将组合键放在一个`tuple`中传入`input()`即可。

```
ele.input((Keys.CTRL, 'a', Keys.DEL))  # ctrl+a+del
```

`Keys`内置了 5 个常用组合键，分别为`CTRL_A`、`CTRL_C`、`CTRL_X`、`CTRL_V`、`CTRL_Z`、`CTRL_Y`。

```
ele.input(Keys.CTRL_A)  # 全选
```

---

### 📌 `focus()`[​](#-focus "-focus的直接链接")

此方法用于使元素获取焦点。

**参数：** 无

**返回：** `None`

---

✅️️ 拖拽和悬停[​](#️️-拖拽和悬停 "✅️️ 拖拽和悬停的直接链接")
----------------------------------------

Tips

除了以下方法，本库还提供更灵活的动作链功能，详见后面章节。

### 📌 `drag()`[​](#-drag "-drag的直接链接")

此方法用于拖拽元素到相对于当前的一个新位置，可以设置速度。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `offset_x` | `int` | `0` | x 轴偏移量，向下向右为正 |
| `offset_y` | `int` | `0` | y 轴偏移量，向下向右为正 |
| `duration` | `float` | `0.5` | 用时，单位秒，传入`0`即瞬间到达 |

**返回：**`None`

**示例：**

```
# 拖动当前元素到距离50*50的位置，用时1秒  
ele.drag(50, 50, 1)
```

---

### 📌 `drag_to()`[​](#-drag_to "-drag_to的直接链接")

此方法用于拖拽元素到另一个元素上或一个坐标上。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `ele_or_loc` | `ChromiumElement` `Tuple[int, int]` | 必填 | 另一个元素对象或坐标元组 |
| `duration` | `float` | `0.5` | 用时，单位秒，传入`0`即瞬间到达 |

**返回：**`None`

**示例：**

```
# 把 ele1 拖拽到 ele2 上  
ele1 = page.ele('#div1')  
ele2 = page.ele('#div2')  
ele1.drag_to(ele2)  
  
# 把 ele1 拖拽到网页 50, 50 的位置  
ele1.drag_to((50, 50))
```

---

### 📌 `hover()`[​](#-hover "-hover的直接链接")

此方法用于模拟鼠标悬停在元素上，可接受偏移量，偏移量相对于元素左上角坐标。不传入`offset_x`和`offset_y`值时悬停在元素中点。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `offset_x` | `int` | `None` | 相对元素左上角坐标的 x 轴偏移量，向下向右为正 |
| `offset_y` | `int` | `None` | 相对元素左上角坐标的 y 轴偏移量，向下向右为正 |

**返回：**`None`

**示例：**

```
# 悬停在元素右上方 50*50 的位置  
ele.hover(50, -50)  
  
# 悬停在元素上中部，x 相对左上角向右偏移50，y 保持在元素中点  
ele.hover(offset_x=50)  
  
# 悬停在元素中点  
ele.hover()
```

---

  

✅️️ 修改元素[​](#️️-修改元素 "✅️️ 修改元素的直接链接")
-------------------------------------

### 📌 `set.innerHTML()`[​](#-setinnerhtml "-setinnerhtml的直接链接")

此方法用于设置元素的 innerHTML 内容。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `html` | `str` | 必填 | html文本 |

**返回：**`None`

---

### 📌 `set.property()`[​](#-setproperty "-setproperty的直接链接")

此方法用于设置元素`property`属性。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `name` | `str` | 必填 | 属性名 |
| `value` | `str` | 必填 | 属性值 |

**返回：**`None`

**示例：**

```
ele.set.property('value', 'Hello world!')
```

---

### 📌 `set.style()`[​](#-setstyle "-setstyle的直接链接")

此方法用于设置元素样式。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `name` | `str` | 必填 | 属性名 |
| `value` | `str` | 必填 | 属性值 |

**返回：**`None`

---

### 📌 `set.attr()`[​](#-setattr "-setattr的直接链接")

此方法用于设置元素 attribute 属性。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `name` | `str` | 必填 | 属性名 |
| `value` | `str` | 必填 | 属性值 |

**返回：**`None`

**示例：**

```
ele.set.attr('href', 'http://DrissionPage.cn')
```

---

### 📌 `remove_attr()`[​](#-remove_attr "-remove_attr的直接链接")

此方法用于删除元素 attribute 属性。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `name` | `str` | 必填 | 属性名 |

**返回：**`None`

**示例：**

```
ele.remove_attr('href')
```

---

### 📌 `set.value()`[​](#-setvalue "-setvalue的直接链接")

此方法用于设置元素`value`值。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `value` | `str` | 必填 | 属性值 |

**返回：**`None`

---

### 📌 `check()`[​](#-check "-check的直接链接")

此方法用于选中或取消选中元素。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `uncheck` | `bool` | `False` | 是否取消选中 |
| `by_js` | `bool` | `False` | 是否用 js 方式选择 |

**返回：**`None`

---

✅️️ 执行 js 脚本[​](#️️-执行-js-脚本 "✅️️ 执行 js 脚本的直接链接")
-------------------------------------------------

### 📌 `run_js()`[​](#-run_js "-run_js的直接链接")

此方法用于对元素执行 js 代码，代码中用`this`表示元素自己。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `script` | `str` | 必填 | js 脚本文本或脚本文件路径 |
| `*args` | - | 无 | 传入的参数，按顺序在js文本中对应`arguments[0]`、`arguments[1]`... |
| `as_expr` | `bool` | `False` | 是否作为表达式运行，为`True`时`args`参数无效 |
| `timetout` | `float` | `None` | js 超时时间（秒），为`None`则使用页面`timeouts.script`设置 |

| 返回类型 | 说明 |
| --- | --- |
| `Any` | 脚本执行结果 |

注意

要获取 js 结果记得写上`return`。

**示例：**

```
# 用执行 js 的方式点击元素  
ele.run_js('this.click();')  
  
# 用 js 获取元素高度  
height = ele.run_js('return this.offsetHeight;')
```

---

### 📌 `run_async_js()`[​](#-run_async_js "-run_async_js的直接链接")

此方法用于以异步方式执行 js 代码，代码中用`this`表示元素自己。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `script` | `str` | 必填 | js 脚本文本 |
| `*args` | - | 无 | 传入的参数，按顺序在js文本中对应`arguments[0]`、`arguments[1]`... |
| `as_expr` | `bool` | `False` | 是否作为表达式运行，为`True`时`args`参数无效 |

**返回：**`None`

---

### 📌 `add_init_js()`[​](#-add_init_js "-add_init_js的直接链接")

此方法用于添加初始化脚本，在页面加载任何脚本前执行。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `script` | `str` | 必填 | js 脚本文本 |

| 返回类型 | 说明 |
| --- | --- |
| `str` | 添加的脚本的 id |

---

### 📌 `remove_init_js()`[​](#-remove_init_js "-remove_init_js的直接链接")

此方法用于删除初始化脚本，`script_id`传入`None`时删除所有。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `script_id` | `str` | `None` | 脚本的id，传入`None`时删除所有 |

**返回：**`None`

---

✅️️ 元素滚动[​](#️️-元素滚动 "✅️️ 元素滚动的直接链接")
-------------------------------------

元素滚动功能藏在`scroll`属性中。用于使可滚动的容器元素内部进行滚动，或使元素本身滚动到可见。

```
# 滚动到底部  
ele.scroll.to_bottom()  
  
# 滚动到最右边  
ele.scroll.to_rightmost()  
  
# 向下滚动 200 像素  
ele.scroll.down(200)  
  
# 滚动到指定位置  
ele.scroll.to_location(100, 300)  
  
# 滚动页面使自己可见  
ele.scroll.to_see()
```

### 📌 `scroll()`或`scroll.down()`[​](#-scroll或scrolldown "-scroll或scrolldown的直接链接")

这两个方法效果是一样的，用于使元素向下滚动若干像素，水平位置不变。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `pixel` | `int` | 必填 | 滚动的像素 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 执  行滚动的元素自身 |

```
# 向下滚动30像素  
ele.scroll(30)  
ele.scroll.down(30)
```

---

### 📌 `scroll.up()`[​](#-scrollup "-scrollup的直接链接")

此方法用于使元素向上滚动若干像素，水平位置不变。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `pixel` | `int` | 必填 | 滚动的像素 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 执行滚动的元素自身 |

**示例：**

```
page.scroll.up(30)
```

---

### 📌 `scroll.right()`[​](#-scrollright "-scrollright的直接链接")

此方法用于使元素内滚动条向右滚动若干像素，垂直位置不变。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `pixel` | `int` | 必填 | 滚动的像素 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 执行滚动的元素自身 |

---

### 📌 `scroll.left()`[​](#-scrollleft "-scrollleft的直接链接")

此方法用于使元素内滚动条向左滚动若干像素，垂直位置不变。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `pixel` | `int` | 必填 | 滚动的像素 |

|   返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 执行滚动的元素自身 |

---

### 📌 `scroll.to_top()`[​](#-scrollto_top "-scrollto_top的直接链接")

此方法用于滚动到元素顶部，水平位置不变。

**参数：** 无

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 执行滚动的元素自身 |

**示例：**

```
page.scroll.to_top()
```

---

### 📌 `scroll.to_bottom()`[​](#-scrollto_bottom "-scrollto_bottom的直接链接")

此方法用于滚动到元素底部，水平位置不变。

**参数：** 无

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 执行滚动的元素自身 |

---

### 📌 `scroll.to_half()`[​](#-scrollto_half "-scrollto_half的直接链接")

此方法用于滚动到元素垂直中间位置，水平位置不变。

**参数：** 无

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 执行滚动的元素自身 |

---

### 📌 `scroll.to_rightmost()`[​](#-scrollto_rightmost "-scrollto_rightmost的直接链接")

此方法用于滚动到元素最右边，垂直位置不变。

**参数：** 无

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 执行滚动的元素自身 |

---

### 📌 `scroll.to_leftmost()`[​](#-scrollto_leftmost "-scrollto_leftmost的直接链接")

此方法用于滚动到元素最左边，垂直位置不变。

**参数：** 无

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 执行滚动的元素自身 |

---

### 📌 `scroll.to_location()`[​](#-scrollto_location "-scrollto_location的直接链接")

此方法用于滚动到元素滚动到指定位置。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `x` | `int` | 必填 | 水平位置 |
| `y` | `int` | 必填 | 垂直位置 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 执行滚动的元素自身 |

**示例：**

```
page.scroll.to_location(300, 50)
```

---

### 📌 `scroll.to_see()`[​](#-scrollto_see "-scrollto_see的直接链接")

此方法用于滚动页面直到元素可见。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `center` | `bool` `None` | `None` | 是否尽量滚动到页面正中，为`None`时如果被遮挡，则滚动到页面正中 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 执行滚动的元素自身 |

---

### 📌 `scroll.to_center()`[​](#-scrollto_center "-scrollto_center的直接链接")

此方法用于尽量把元素滚动到视口正中。

**参数：** 无

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 执行滚动的元素自身 |

---

✅️️ 列表选择[​](#️️-列表选择 "✅️️ 列表选择的直接链接")
-------------------------------------

`<select>`下拉列表元素功能在`select`属性中。可自动等待列表项出现再实施选择。

此属性用于对`<select>`元素的操作。非`<select>`元素此属性为`None`。

假设有以下`<select>`元素，下面示例以此为基础：

```
<select id='s' multiple>  
    <option value='value1'>text1</option>  
    <option value='value2'>text2</option>  
    <option value='value3'>text3</option>  
</select>
```

### 📌 点击列表项元素进行选取[​](#-点击列表项元素进行选取 "📌 点击列表项元素进行选取的直接链接")

可以获取`<option>`元素，进行选取或取消选择。

**示例：**

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
ele = tab('t:select')('t:option')  
ele.click()
```

---

### 📌 `select()`和`select.by_text()`[​](#-select和selectby_text "-select和selectby_text的直接链接")

这两个方法功能一样，用于按文本选择列表项。如为多选列表，可多选。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `text` | `str` `list` `tuple` | 必填 | 作为选择条件的文本，传入`list`或`tuple`可选择多项 |
| `timeout` | `float` | `None` | 超时时间（秒），为`None`默认使用页面超时时间 |

| 返回类型 | 说明 |
| --- | --- |
| `bool` | 是否选择成功 |

---

### 📌 `select.by_value()`[​](#-selectby_value "-selectby_value的直接链接")

此方法用于按`value`属性选择列表项。如为多选列表，可多选。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `value` | `str` `list` `tuple` | 必填 | 作为选择条件的`value`值，传入`list`或`tuple`可选择多项 |
| `timeout` | `float` | `None` | 超时时间（秒），为`None`默认使用页面超时时间 |

| 返回类型 | 说明 |
| --- | --- |
| `bool` | 是否选择成功 |

---

### 📌 `select.by_index()`[​](#-selectby_index "-selectby_index的直接链接")

此方法用于按序号选择列表项，从`1`开始。如为多选列表，可多选。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `index` | `int` `list` `tuple` | 必填 | 选择第几项，传入`list`或`tuple`可选择多项 |
| `timeout` | `float` | `None` | 超时时间（秒），为`None`默认使用页面超时时间 |

| 返回类型 | 说明 |
| --- | --- |
| `bool` | 是否选择成功 |

---

### 📌 `select.by_locator()`[​](#-selectby_locator "-selectby_locator的直接链接")

此方法可用定位符筛选选项元素。如为多选列表，可多选。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `list` `tuple` | 必填 | 定位符，传入`list`或`tuple`可选择多项 |
| `timeout` | `float` | `None` | 超时时间（秒），为`None`默认使用页面超时时间 |

| 返回类型 | 说明 |
| --- | --- |
| `bool` | 是否选择成功 |

---

### 📌 `select.by_option()`[​](#-selectby_option "-selectby_option的直接链接")

此方法用于选中单个或多个列表项元素。如为多选列表，可多选。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `option` | `ChromiumElement` `List[ChromiumElement]` | 必填 | `<option>`元素或它们组成的列表 |

**返回：**`None`

**示例：**

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
select = tab('t:select')  
option = select('t:option')  
select.select.by_option(option)
```

---

### 📌 `select.cancel_by_text()`[​](#-selectcancel_by_text "-selectcancel_by_text的直接链接")

此方法用于按文本取消选择列表项。如为多选列表，可取消多项。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `text` | `str` `list` `tuple` | 必填 | 作为选择条件的文本，传入`list`或`tuple`可选择多项 |
| `timeout` | `float` | `None` | 超时时间（秒），为`None`默认使用页面超时时间 |

| 返回类型 | 说明 |
| --- | --- |
| `bool` | 是否选择成功 |

---

### 📌 `select.cancel_by_value()`[​](#-selectcancel_by_value "-selectcancel_by_value的直接链接")

此方法用于按`value`属性取消选择列表项。如为多选列表，可取消多项。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `value` | `str` `list` `tuple` | 必填 | 作为选择条件的`value`值，传入`list`或`tuple`可选择多项 |
| `timeout` | `float` | `None` | 超时时间（秒），为`None`默认使用页面超时时间 |

| 返回类型 | 说明 |
| --- | --- |
| `bool` | 是否选择成功 |

---

### 📌 `select.cancel_by_index()`[​](#-selectcancel_by_index "-selectcancel_by_index的直接链接")

此方法用于按序号取消选择列表项，从`1`开始。如为多选列表，可取消多项。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `index` | `int` `list` `tuple` | 必填 | 选择第几项，传入`list`或`tuple`可选择多项 |
| `timeout` | `float` | `None` | 超时时间（秒），为`None`默认使用页面超时时间 |

| 返回类型 | 说明 |
| --- | --- |
| `bool` | 是否选择成功 |

---

### 📌 `select.cancel_by_locator()`[​](#-selectcancel_by_locator "-selectcancel_by_locator的直接链接")

此方法可用定位符筛选选项元素。如为多选列表，可取消多项。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `list` `tuple` | 必填 | 定位符，传入`list`或`tuple`可选择多项 |
| `timeout` | `float` | `None` | 超时时间，为`None`默认使用页面超时时间 |

| 返回类型 | 说明 |
| --- | --- |
| `bool` | 是否选择成功 |

---

### 📌 `select.cancel_by_option()`[​](#-selectcancel_by_option "-selectcancel_by_option的直接链接")

此方法用于取消选中单个或多个列表项元素。如为多选列表，可多选。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `option` | `ChromiumElement` `List[ChromiumElement]` | 必填 | `<option>`元素或它们组成的列表 |

**返回：**`None`

---

### 📌 `select.all()`[​](#-selectall "-selectall的直接链接")

此方法用于全选所有项。多选列表才有效。

**参数：** 无

**返回：**`None`

---

### 📌 `select.clear()`[​](#-selectclear "-selectclear的直接链接")

此方法用于取消所有项选中状态。多选列表才有效。

**参数：** 无

**返回：**`None`

---

### 📌 `select.invert()`[​](#-selectinvert "-selectinvert的直接链接")

此方法用于反选。多选列表才有效。

**参数：** 无

**返回：**`None`

---

### 📌 `select.is_multi`[​](#-selectis_multi "-selectis_multi的直接链接")

此属性返回当前元素是否多选列表。

**返回类型：**`bool`

---

### 📌 `select.options`[​](#-selectoptions "-selectoptions的直接链接")

此属性返回当前列表元素所有选项元素对象。

**返回类型：**`ChromiumElement`

---

### 📌 `select.selected_option`[​](#-selectselected_option "-selectselected_option的直接链接")

此属性返回当前元素选中的选项（单选列表）。

**返回类型：**`ChromiumElement`

---

### 📌 `select.selected_options`[​](#-selectselected_options "-selectselected_options的直接链接")

此属性返回当前元素所有选中的选项（多选列表）。

**返回类型：**`List[ChromiumElement]`

[上一页

🔦 语法速查表](/browser_control/get_elements/sheet)[下一页

🛰️ 获取元素信息](/browser_control/get_ele_info)

* [✅️️ 点击元素](#️️-点击元素)
  + [📌 `click()`和`click.left()`](#-click和clickleft)
  + [📌 `click.right()`](#-clickright)
  + [📌 `click.middle()`](#-clickmiddle)
  + [📌 `click.multi()`](#-clickmulti)
  + [📌 `click.at()`](#-clickat)
  + [📌 `click.to_upload()`](#-clickto_upload)
  + [📌 `click.to_download()`](#-clickto_download)
  + [📌 `click.for_new_tab()`](#-clickfor_new_tab)
* [✅️️ 输入内容](#️️-输入内容)
  + [📌 `clear()`](#-clear)
  + [📌 `input()`](#-input)
  + [📌 输入组合键](#-输入组合键)
  + [📌 `focus()`](#-focus)
* [✅️️ 拖拽和悬停](#️️-拖拽和悬停)
  + [📌 `drag()`](#-drag)
  + [📌 `drag_to()`](#-drag_to)
  + [📌 `hover()`](#-hover)
* [✅️️ 修改元素](#️️-修改元素)
  + [📌 `set.innerHTML()`](#-setinnerhtml)
  + [📌 `set.property()`](#-setproperty)
  + [📌 `set.style()`](#-setstyle)
  + [📌 `set.attr()`](#-setattr)
  + [📌 `remove_attr()`](#-remove_attr)
  + [📌 `set.value()`](#-setvalue)
  + [📌 `check()`](#-check)
* [✅️️ 执行 js 脚本](#️️-执行-js-脚本)
  + [📌 `run_js()`](#-run_js)
  + [📌 `run_async_js()`](#-run_async_js)
  + [📌 `add_init_js()`](#-add_init_js)
  + [📌 `remove_init_js()`](#-remove_init_js)
* [✅️️ 元素滚动](#️️-元素滚动)
  + [📌 `scroll()`或`scroll.down()`](#-scroll或scrolldown)
  + [📌 `scroll.up()`](#-scrollup)
  + [📌 `scroll.right()`](#-scrollright)
  + [📌 `scroll.left()`](#-scrollleft)
  + [📌 `scroll.to_top()`](#-scrollto_top)
  + [📌 `scroll.to_bottom()`](#-scrollto_bottom)
  + [📌 `scroll.to_half()`](#-scrollto_half)
  + [📌 `scroll.to_rightmost()`](#-scrollto_rightmost)
  + [📌 `scroll.to_leftmost()`](#-scrollto_leftmost)
  + [📌 `scroll.to_location()`](#-scrollto_location)
  + [📌 `scroll.to_see()`](#-scrollto_see)
  + [📌 `scroll.to_center()`](#-scrollto_center)
* [✅️️ 列表选择](#️️-列表选择)
  + [📌 点击列表项元素进行选取](#-点击列表项元素进行选取)
  + [📌 `select()`和`select.by_text()`](#-select和selectby_text)
  + [📌 `select.by_value()`](#-selectby_value)
  + [📌 `select.by_index()`](#-selectby_index)
  + [📌 `select.by_locator()`](#-selectby_locator)
  + [📌 `select.by_option()`](#-selectby_option)
  + [📌 `select.cancel_by_text()`](#-selectcancel_by_text)
  + [📌 `select.cancel_by_value()`](#-selectcancel_by_value)
  + [📌 `select.cancel_by_index()`](#-selectcancel_by_index)
  + [📌 `select.cancel_by_locator()`](#-selectcancel_by_locator)
  + [📌 `select.cancel_by_option()`](#-selectcancel_by_option)
  + [📌 `select.all()`](#-selectall)
  + [📌 `select.clear()`](#-selectclear)
  + [📌 `select.invert()`](#-selectinvert)
  + [📌 `select.is_multi`](#-selectis_multi)
  + [📌 `select.options`](#-selectoptions)
  + [📌 `select.selected_option`](#-selectselected_option)
  + [📌 `select.selected_options`](#-selectselected_options)

* 🚀 控制浏览器
* 🛰️ 获取元素信息

本页总览

🛰️ 获取元素信息
=========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

浏览器元素对应的对象是`ChromiumElement`和`ShadowRoot`，本节介绍如何获取元素信息。

✅️️ 内容和属性[​](#️️-内容和属性 "✅️️ 内容和属性的直接链接")
----------------------------------------

### 📌 `tag`[​](#-tag "-tag的直接链接")

此属性返回元素的标签名。

**返回类型：**`str`

---

### 📌 `html`[​](#-html "-html的直接链接")

此属性返回元素的`outerHTML`文本。

**返回类型：**`str`

---

### 📌 `inner_html`[​](#-inner_html "-inner_html的直接链接")

此属性返回元素的`innerHTML`文本。

**返回类型：**`str`

---

### 📌 `text`[​](#-text "-text的直接链接")

此属性返回元素内所有文本组合成的字符串。  
该字符串已格式化，即已转码，已去除多余换行符，符合人读取习惯，便于直接使用。

**返回类型：**`str`

---

### 📌 `raw_text`[​](#-raw_text "-raw_text的直接链接")

此属性返回元素内未经处理的原始文本。

**返回类型：**`str`

---

### 📌 `texts()`[​](#-texts "-texts的直接链接")

此方法返回元素内所有**直接**子节点的文本，包括元素和文本节点。 它有一个参数`text_node_only`，为`True`时则只获取只返回不被包裹的文本节点。这个方法适用于获取文本节点和元素节点混排的情况。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `text_node_only` | `bool` | `False` | 是否只返回文本节点 |

| 返回类型 | 说明 |
| --- | --- |
| `List[str]` | 文本列表 |

---

### 📌 `comments`[​](#-comments "-comments的直接链接")

此属性以列表形式返回元素内的注释。

**返回类型：**`List[str]`

---

### 📌 `attrs`[​](#-attrs "-attrs的直接链接")

此属性以字典形式返回元素所有属性及值。

**返回类型：**`dict`

---

### 📌 `attr()`[​](#-attr "-attr的直接链接")

此方法返回元素某个 attribute 属性值。它接收一个字符串参数，返回该属性值文本，无该属性时返回`None`。  
此属性返回的`src`、`href`属性为已补充完整的路径。`text`属性为已格式化文本。
如果要获取未补充完整路径的`src`或`href`属性，可以用`attrs['src']`。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `name` | `str` | 必填 | 属性名称 |

| 返回类型 | 说明 |
| --- | --- |
| `str` | 属性值文本 |
| `None` | 没有该属性返回`None` |

---

### 📌 `property()`[​](#-property "-property的直接链接")

此方法返回`property`属性值。它接收一个字符串参数，返回该参数的属性值。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `name` | `str` | 必填 | 属性名称 |

| 返回类型 | 说明 |
| --- | --- |
| `str` | 属性值 |

---

### 📌 `value`[​](#-value "-value的直接链接")

此方法返回元素的`value`值。

**返回类型：**`str`

---

### 📌 `link`[​](#-link "-link的直接链接")

此方法返回元素的`href`属性或`src`属性，没有这两个属性则返回`None`。

**返回类型：**`str`

---

### 📌 `pseudo.before`[​](#-pseudobefore "-pseudobefore的直接链接")

此属性以文本形式返回当前元素的`::before`伪元素内容。

**类型：**`str`

---

### 📌 `pseudo.after`[​](#-pseudoafter "-pseudoafter的直接链接")

此属性以文本形式返回当前元素的`::after`伪元素内容。

**类型：**`str`

---

### 📌 `style()`[​](#-style "-style的直接链接")

该方法返回元素 css 样式属性值，可获取伪元素的属性。 它有两个参数，`style`参数输入样式属性名称，`pseudo_ele`
参数输入伪元素名称，省略则获取普通元素的 css 样式属性。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `style` | `str` | 必填 | 样式名称 |
| `pseudo_ele` | `str` | `''` | 伪元素名称（如有） |

| 返回类型 | 说明 |
| --- | --- |
| `str` | 样式属性值 |

---

### 📌 `shadow_root`[​](#-shadow_root "-shadow_root的直接链接")

此属性返回元素内的 shadow-root 对象，没有的返回`None`。

**类型：**`ShadowRoot`

---

### 📌 `child_count`[​](#-child_count "-child_count的直接链接")

此属性返回元素内第一级子元素个数。

**类型：**`int`

---

✅️️ 大小和位置[​](#️️-大小和位置 "✅️️ 大小和位置的直接链接")
----------------------------------------

### 📌 `rect.size`[​](#-rectsize "-rectsize的直接链接")

此属性以元组形式返回元素的大小。

**类型：**`Tuple[float, float]`

```
size = ele.rect.size  
# 返回：(50, 50)
```

---

### 📌 `rect.location`[​](#-rectlocation "-rectlocation的直接链接")

此属性以元组形式返回元素**左上角**在**整个页面**中的坐标。

**类型：**`Tuple[float, float]`

```
loc = ele.rect.location  
# 返回：(50, 50)
```

---

### 📌 `rect.midpoint`[​](#-rectmidpoint "-rectmidpoint的直接链接")

此属性以元组形式返回元素**中点**在**整个页面**中的坐标。

**类型：**`Tuple[float, float]`

```
loc = ele.rect.midpoint  
# 返回：(55, 55)
```

---

### 📌 `rect.click_point`[​](#-rectclick_point "-rectclick_point的直接链接")

此属性以元组形式返回元素**点击点**在**整个页面**中的坐标。

点击点是指`click()`方法点击时的位置，位于元素中上部。

**类型：**`Tuple[float, float]`

---

### 📌 `rect.corners`[​](#-rectcorners "-rectcorners的直接链接")

此属性以列表形式返回元素四个角在页面中的坐标，顺序：左上、右上、右下、左下。

**类型：**`((float, float), (float, float), (float, float), (float, float),)`

---

### 📌 `rect.viewport_corners`[​](#-rectviewport_corners "-rectviewport_corners的直接链接")

此属性以列表形式返回元素四个角在视口中的坐标，顺序：左上、右上、右下、左下。

**类型：**`list[(float, float), (float, float), (float, float), (float, float)]`

---

### 📌 `rect.viewport_location`[​](#-rectviewport_location "-rectviewport_location的直接链接")

此属性以元组形式返回元素**左上角**在**当前视口**中的坐标。

**类型：**`Tuple[float, float]`

---

### 📌 `rect.viewport_midpoint`[​](#-rectviewport_midpoint "-rectviewport_midpoint的直接链接")

此属性以元组形式返回元素**中点**在**当前视口**中的坐标。

**类型：**`Tuple[floatt, float]`

---

### 📌 `rect.viewport_click_point`[​](#-rectviewport_click_point "-rectviewport_click_point的直接链接")

此属性以元组形式返回元素**点击点**在**当前视口**中的坐标。

**类型：**`Tuple[float, float]`

---

### 📌 `rect.screen_location`[​](#-rectscreen_location "-rectscreen_location的直接链接")

此属性以元组形式返回元素**左上角**在**屏幕**中的坐标。

**类型：**`Tuple[float, float]`

---

### 📌 `rect.screen_midpoint`[​](#-rectscreen_midpoint "-rectscreen_midpoint的直接链接")

此属性以元组形式返回元素**中点**在**屏幕**中的坐标。

**类型：**`Tuple[float, float]`

---

### 📌 `rect.screen_click_point`[​](#-rectscreen_click_point "-rectscreen_click_point的直接链接")

此属性以元组形式返回元素**点击点**在**屏幕**中的坐标。

**类型：**`Tuple[float, float]`

---

### 📌 `rect.scroll_position`[​](#-rectscroll_position "-rectscroll_position的直接链接")

此属性返回元素内滚动条位置，格式：(x, y)。

**类型：**`Tuple[float, float]`

---

### 📌 `xpath`[​](#-xpath "-xpath的直接链接")

此属性返回当前元素在页面中 xpath 的绝对路径。

**返回类型：**`str`

---

### 📌 `css_path`[​](#-css_path "-css_path的直接链接")

此属性返回当前元素在页面中 css selector 的绝对路径。

**返回类型：**`str`

---

✅️️ 元素列表中批量获取信息[​](#️️-元素列表中批量获取信息 "✅️️ 元素列表中批量获取信息的直接链接")
----------------------------------------------------------

`eles()`等返回的元素列表，自带`get`属性，可用于获取指定信息。

### 📌 示例[​](#-示例 "📌 示例的直接链接")

```
from DrissionPage import SessionPage  
  
page = SessionPage()  
page.get('https://www.baidu.com')  
eles = page('#s-top-left').eles('t:a')  
print(eles.get.texts())  # 获取所有元素的文本
```

**输出：**

```
['新闻', 'hao123', '地图', '贴吧', '视频', '图片', '网盘', '文库', '更多', '翻译', '学术', '百科', '知道', '健康', '营销推广', '直播', '音乐', '橙篇', '查看全部百度产品 >']
```

### 📌 `get.attrs()`[​](#-getattrs "-getattrs的直接链接")

此方法用于返回所有元素指定的 attribute 属性组成的列表。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `name` | `str` | 必填 | 属性名称 |

| 返回类型 | 说明 |
| --- | --- |
| `List[str]` | 属性文本组成的列表 |

---

### 📌 `get.links()`[​](#-getlinks "-getlinks的直接链接")

此方法用于返回所有元素的`link`属性组成的列表。

**参数：** 无

| 返回类型 | 说明 |
| --- | --- |
| `List[str]` | 链接文本组成的列表 |

---

### 📌 `get.texts()`[​](#-gettexts "-gettexts的直接链接")

此方法用于返回所有元素的`text`属性组成的列表。

**参数：** 无

| 返回类型 | 说明 |
| --- | --- |
| `List[str]` | 元素文本组成的列表 |

---

  

✅️️ 状态信息[​](#️️-状态信息 "✅️️ 状态信息的直接链接")
-------------------------------------

### 📌`timeout`[​](#timeout "timeout的直接链接")

此属性返回获取内部或相对定位元素的超时时间，实际上是元素所在页面的超时设置。

**类型：**`float`

---

### 📌`states.is_in_viewport`[​](#statesis_in_viewport "statesis_in_viewport的直接链接")

此属性以布尔值方式返回元素是否在视口中，以元素可以接受点击的点为判断。

**类型：**`bool`

---

### 📌`states.is_whole_in_viewport`[​](#statesis_whole_in_viewport "statesis_whole_in_viewport的直接链接")

此属性以布 尔值方式返回元素是否整个在视口中。

**类型：**`bool`

---

### 📌`states.is_alive`[​](#statesis_alive "statesis_alive的直接链接")

此属性以布尔值形式返回当前元素是否仍可用。用于判断 d 模式下是否因页面刷新而导致元素失效。

**类型：**`bool`

---

### 📌 `states.is_checked`[​](#-statesis_checked "-statesis_checked的直接链接")

此属性以布尔值返回表单单选或多选元素是否选中。

**类型：**`bool`

---

### 📌 `states.is_selected`[​](#-statesis_selected "-statesis_selected的直接链接")

此属性以布尔值返回`<select>`元素中的项是否选中。

**类型：**`bool`

---

### 📌 `states.is_enabled`[​](#-statesis_enabled "-statesis_enabled的直接链接")

此属性以布尔值返回元素是否可用。

**类型：**`bool`

---

### 📌 `states.is_displayed`[​](#-statesis_displayed "-statesis_displayed的直接链接")

此属性以布尔值返回元素是否可见。

**类型：**`bool`

---

### 📌 `states.is_covered`[​](#-statesis_covered "-statesis_covered的直接链接")

此属性返回元素是否被其它元素覆盖。如被覆盖，返回覆盖元素的 id，否则返回`False`

| 返回类型 | 说明 |
| --- | --- |
| `False` | 未被覆盖，返回`False` |
| `int` | 被覆盖时返回覆盖元素的 id |

---

### 📌 `states.is_clickable`[​](#-statesis_clickable "-statesis_clickable的直接链接")

此属性返回元素是否可被模拟点击，从是否有大小、是否可用、是否显示、是否响应点击判断，不判断是否被遮挡。

**类型：**`bool`

---

### 📌 `states.has_rect`[​](#-stateshas_rect "-stateshas_rect的直接链接")

此属性返回元素是否拥有大小和位置信息，有则返回四个角在页面上的坐标组成的列表，没有则返回`False`。

| 返回类型 | 说明 |
| --- | --- |
| `list` | 存在大小和位置信息时，以[(int, int), ...] 格式返回元素四个角的坐标，顺序：左上、右上、右下、左下 |
| `False` | 不存在时返回`False` |

---

✅️️ 保存元素[​](#️️-保存元素 "✅️️ 保存元素的直接链接")
-------------------------------------

保存功能是本库一个特色功能，可以直接读取浏览器缓存，无需依赖另外的 ui 库或重新下载就可以保存页面资源。

作为对比，selenium 无法自身实现图片另存，往往需要通过使用 ui 工具进行辅助，不仅效率和可靠性低，还占用键鼠资源。

### 📌 `src()`[​](#-src "-src的直接链接")

此方法用于返回元素`src`属性所使用的资源。base64 的可转为`bytes`返回，其它的以`str`返回。无资源的返回`None`。

例如，可获取页面上图片字节数据，用于识别内容，或保存到文件。`<script>`和`<link>`标签也可获取文件内容。

注意

获取`<script>`或`<link>`文件内容时，视网站情况不一定会成功。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` | `None` | 等待资源加载超时时间（秒），为`None`时使用元素所在页面`timeout`属性 |
| `base64_to_bytes` | `bool` | `True` | 为`True`时，如果是 base64 数据，转换为`bytes`格式 |

| 返回类型 | 说明 |
| --- | --- |
| `str` | 资源字符串 |
| `None` | 无资源的返回`None` |

**示例：**

```
img = page('tag:img')  
src = img.src()
```

---

### 📌 `save()`[​](#-save "-save的直接链接")

此方法用于保存`src()`方法获取到的资源到文件。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `path` | `str` `Path` | `None` | 文件保存路径，为`None`时保存到当前文件夹 |
| `name` | `str` | `None` | 文件名称，需包含后缀，为`None`时从资源 url 获取 |
| `timeout` | `float` | `None` | 等待资源加载超时时间（秒），为`None`时使用元素所在页面`timeout`属性 |
| `rename` | `bool` | `True` | 遇到重名文件时是否自动重命名 |

| 返回类型 | 说明 |
| --- | --- |
| `str` | 保存路径 |

**示例：**

```
img = page('tag:img')  
img.save('D:\\img.png')
```

---

✅️️ 比较元素[​](#️️-比较元素 "✅️️ 比较元素的直接链接")
-------------------------------------

两个元素对象可以用`==`来比较，以判断它们是否指向同一个元素。

**示例：**

```
ele1 = page('t:div')  
ele2 = page('t:div')  
print(ele1==ele2)  # 输出True
```

[上一页

🛰️ 元素交互](/browser_control/ele_operation)[下一页

🛰️ iframe 操作](/browser_control/iframe)

* [✅️️ 内容和属性](#️️-内容和属性)
  + [📌 `tag`](#-tag)
  + [📌 `html`](#-html)
  + [📌 `inner_html`](#-inner_html)
  + [📌 `text`](#-text)
  + [📌 `raw_text`](#-raw_text)
  + [📌 `texts()`](#-texts)
  + [📌 `comments`](#-comments)
  + [📌 `attrs`](#-attrs)
  + [📌 `attr()`](#-attr)
  + [📌 `property()`](#-property)
  + [📌 `value`](#-value)
  + [📌 `link`](#-link)
  + [📌 `pseudo.before`](#-pseudobefore)
  + [📌 `pseudo.after`](#-pseudoafter)
  + [📌 `style()`](#-style)
  + [📌 `shadow_root`](#-shadow_root)
  + [📌 `child_count`](#-child_count)
* [✅️️ 大小和位置](#️️-大小和位置)
  + [📌 `rect.size`](#-rectsize)
  + [📌 `rect.location`](#-rectlocation)
  + [📌 `rect.midpoint`](#-rectmidpoint)
  + [📌 `rect.click_point`](#-rectclick_point)
  + [📌 `rect.corners`](#-rectcorners)
  + [📌 `rect.viewport_corners`](#-rectviewport_corners)
  + [📌 `rect.viewport_location`](#-rectviewport_location)
  + [📌 `rect.viewport_midpoint`](#-rectviewport_midpoint)
  + [📌 `rect.viewport_click_point`](#-rectviewport_click_point)
  + [📌 `rect.screen_location`](#-rectscreen_location)
  + [📌 `rect.screen_midpoint`](#-rectscreen_midpoint)
  + [📌 `rect.screen_click_point`](#-rectscreen_click_point)
  + [📌 `rect.scroll_position`](#-rectscroll_position)
  + [📌 `xpath`](#-xpath)
  + [📌 `css_path`](#-css_path)
* [✅️️ 元素列表中批量获取信息](#️️-元素列表中批量获取信息)
  + [📌 示例](#-示例)
  + [📌 `get.attrs()`](#-getattrs)
  + [📌 `get.links()`](#-getlinks)
  + [📌 `get.texts()`](#-gettexts)
* [✅️️ 状态信息](#️️-状态信息)
  + [📌`timeout`](#timeout)
  + [📌`states.is_in_viewport`](#statesis_in_viewport)
  + [📌`states.is_whole_in_viewport`](#statesis_whole_in_viewport)
  + [📌`states.is_alive`](#statesis_alive)
  + [📌 `states.is_checked`](#-statesis_checked)
  + [📌 `states.is_selected`](#-statesis_selected)
  + [📌 `states.is_enabled`](#-statesis_enabled)
  + [📌 `states.is_displayed`](#-statesis_displayed)
  + [📌 `states.is_covered`](#-statesis_covered)
  + [📌 `states.is_clickable`](#-statesis_clickable)
  + [📌 `states.has_rect`](#-stateshas_rect)
* [✅️️ 保存元素](#️️-保存元素)
  + [📌 `src()`](#-src)
  + [📌 `save()`](#-save)
* [✅️️ 比较元素](#️️-比较元素)

* 🚀 控制浏览器
* 🔎 查找元素
* 🔦 行为模式

本页总览

🔦 行为模式
======

✅️ 等待元素[​](#️-等待元素 "✅️ 等待元素的直接链接")
----------------------------------

由于网络、js 运行时间的不确定性，经常需要等待元素加载到 DOM 中才能使用。

浏览器所有查找元素操作都自带等待，时间默认跟随元素所在页面`timeout`属性（默认 10 秒），也可以在每次查找时单独设置，单独设置的等待时间不会改变页面原来设置。

### 📌 内置等待[​](#-内置等待 "📌 内置等待的直接链接")

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
# 设置查找元素超时时间为 5 秒  
tab.set.timeouts(5)  
  
# 使用页面超时时间来查找元素（5 秒）  
ele1 = tab.ele('search text')  
# 为这次查找页面独立设置等待时间（1 秒）  
ele1 = tab.ele('search text', timeout=1)  
# 查找后代元素，使用页面超时时间（5 秒）  
ele2 = ele1.ele('search text')  
# 查找后代元素，使用单独设置的超时时间（1 秒）  
ele2 = ele1.ele('some text', timeout=1)
```

### 📌 主动等待[​](#-主动等待 "📌 主动等待的直接链接")

页面对象的`wait.eles_loaded()`方法，可主动等待指定元素加载到 DOM。

用法详见等待章节。

---

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

✅️ 链式写法[​](#️-链式写法 "✅️ 链式写法的直接链接")
----------------------------------

因为元素对象本身又可以查找对象，所有可实现多级链式操作，可使程序更简洁。

**示例：**

```
ele = tab.ele('#s_fm').ele('#su')
```

其中`ele()`可省略，简化写成：

```
ele = tab('#s_fm')('#su')
```

---

✅️ 找不到元素时[​](#️-找不到元素时 "✅️ 找不到元素时的直接链接")
----------------------------------------

### 📌 默认情况[​](#-默认情况 "📌 默认情况的直接链接")

默认情况下，找不到元素时不会立即抛出异常，而是返回一个`NoneElement`对象。

这个对象用`if`判断表现为`False`，调用其功能会抛出`ElementNotFoundError`异常。

这样可以用`if`判断是否找到元素，也可以用`try`去捕获异常。

查找多个元素找不到时，返回空的`list`。

**示例，用`if`判断：**

```
ele = tab.ele('****')  
  
# 判断是否找到元素  
if ele:  
    print('找到了。')  
  
if not ele:  
    print('没有找到。')
```

**示例，用`try`捕获：**

```
try:  
    ele.click()  
except ElementNotFoundError:  
    print('没有找到。')
```

---

### 📌 立即抛出异常[​](#-立即抛出异常 "📌 立即抛出异常的直接链接")

如果想在找不到元素时立刻抛出异常，可以用以下方法设置。

此设置为全局有效，在项目开始时设置一次即可。

查找多个元素找不到时，依然返回空的`list`。

设置全局变量：

```
from DrissionPage.common import Settings  
  
Settings.set_raise_when_ele_not_found(True)
```

**示例：**

```
from DrissionPage import Chromium  
from DrissionPage.common import Settings  
  
Settings.set_raise_when_ele_not_found(True)  
  
tab = Chromium().latest_tab  
tab.get('https://www.baidu.com')  
ele = tab('#abcd')  # ('#abcd')这个元素不存在
```

输出：

```
DrissionPage.errors.ElementNotFoundError:   
没有找到元素。  
method: ele()  
args: {'locator': '#abcd'}
```

---

### 📌 设置默认返回值[​](#-设置默认返回值 "📌 设置默认返回值的直接链接")

如果查找元素后要获取一个属性，但这个元素不一定存在，或者链式查找其中一个节点找不到，可以设置查找失败时返回的值，而不 是抛出异常，可以简化一些采集逻辑。

使用浏览器页面对象的`set.NoneElement_value()`方法设置该值。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `value` | `Any` | `None` | 将返回的设定值 |
| `on_off` | `bool` | `True` | `bool`表示是否启用 |

**返回：**`None`

**示例**

比如说，遍历页面上一个列表中多个对象，但其中有些元素可能缺失某个子元素，可以这样写：

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.set.NoneElement_value('没找到')  
for li in tab.eles('t:li'):  
    name = li('.name').text  
    age = li('.age').text  
    phone = li('.phone').text
```

这样，假如某个子元素不存在，不会抛出异常，而是返回`'没找到'`这个字符串。

[上一页

🔦 相对定位](/browser_control/get_elements/relative)[下一页

🔦 在结果列表中筛选](/browser_control/get_elements/filter)

* [✅️ 等待元素](#️-等待元素)
  + [📌 内置等待](#-内置等待)
  + [📌 主动等待](#-主动等待)
* [✅️ 链式写法](#️-链式写法)
* [✅️ 找不到元素时](#️-找不到元素时)
  + [📌 默认情况](#-默认情况)
  + [📌 立即抛出异常](#-立即抛出异常)
  + [📌 设置默认返回值](#-设置默认返回值)

* 🚀 控制浏览器
* 🔎 查找元素
* 🔦 在结果列表中筛选

本页总览

🔦 在结果列表中筛选
==========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

本节介绍在元素列表中按需要进行筛选，获取指定元素。

`eles()`、`nexts()`等能够获取多个元素的方法，返回的列表可进行进一步筛选，以获取指定的元素。

说明

浏览器页面对象和`SessionPage`产生的元素列表均有此功能，前者筛选功能比后者多。

**示例1，筛选并返回元素列表：**

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.get('https://www.baidu.com')  
eles = tab('#s-top-left').eles('t:a')  # 获取左上角导航栏内所有<a>元素  
for ele in eles.filter.displayed():  # 筛选出显示的元素列表并逐个打印文本  
    print(ele.text, end=' ')
```

**输出：**

```
新闻 hao123 地图 贴吧 视频 图片 网盘 文库 更多
```

**示例2，筛选并返回第一个元素：**

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.get('https://www.baidu.com')  
eles = tab('#s-top-left').eles('t:a')  # 获取左上角导航栏内所有<a>元素  
print(eles.filter_one.displayed().text)  # 筛选出显示的元素并返回第一个
```

**输出：**

```
新闻
```

  

✅️️ 获取单个匹配元素[​](#️️-获取单个匹配元素 "✅️️ 获取单个匹配元素的直接链接")
-------------------------------------------------

说明

静态元素列表只有`filter_one.attr()`和`filter_one.text()`方法。

### 📌 `filter_one.displayed()`[​](#-filter_onedisplayed "-filter_onedisplayed的直接链接")

此方法用于在元素列表中以是否显示为条件筛选元素，返回第一个结果。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `equal` | `bool` | `True` | 是否匹配显示的元素，`False`匹配不显示的 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 匹配成功返回元素对象 |
| `NoneElement` | 失败返回`NoneElement` |
| 抛出`ElementNotFoundError`异常 | `Settings.raise_when_ele_not_found`为`True`时抛出 |

---

### 📌 `filter_one.checked()`[​](#-filter_onechecked "-filter_onechecked的直接链接")

此方法用于在元素列表中以是否被选中为条件筛选元素，返回第一个结果。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `equal` | `bool` | `True` | 是否匹配被选中的元素，`False`匹配不被选中的 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 匹配成功返回元素对象 |
| `NoneElement` | 失败返回`NoneElement` |
| 抛出`ElementNotFoundError`异常 | `Settings.raise_when_ele_not_found`为`True`时抛出 |

---

### 📌 `filter_one.selected()`[​](#-filter_oneselected "-filter_oneselected的直接链接")

此方法用于在元素列表中以是否被选择为条件筛选元素，返回第一个结果。用于`<select>`元素项目。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `equal` | `bool` | `True` | 是否匹配被选择的元素，`False`匹配不被选择的 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 匹配成功返回元素对象 |
| `NoneElement` | 失败返回`NoneElement` |
| 抛出`ElementNotFoundError`异常 | `Settings.raise_when_ele_not_found`为`True`时抛出 |

---

### 📌 `filter_one.enabled()`[​](#-filter_oneenabled "-filter_oneenabled的直接链接")

此方法用于在元素列表中以是否可用为条件筛选元素，返回第一个结果。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `equal` | `bool` | `True` | 是否匹配可用的元素，`False`匹配 disabled 状态的 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 匹配成功返回元素对象 |
| `NoneElement` | 失败返回`NoneElement` |
| 抛出`ElementNotFoundError`异常 | `Settings.raise_when_ele_not_found`为`True`时抛出 |

---

### 📌 `filter_one.clickable()`[​](#-filter_oneclickable "-filter_oneclickable的直接链接")

此方法用于在元素列表中以是否可点击为条件筛选元素，返回第一个结果。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `equal` | `bool` | `True` | 是否匹配可点击的元素，`False`表示匹配不是可点击的 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 匹配成功返回元素对象 |
| `NoneElement` | 失败返回`NoneElement` |
| 抛出`ElementNotFoundError`异常 | `Settings.raise_when_ele_not_found`为`True`时抛出 |

---

### 📌 `filter_one.have_rect()`[​](#-filter_onehave_rect "-filter_onehave_rect的直接链接")

此方法用于在元素列表中以是否有大小为条件筛选元素，返回第一个结果。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `equal` | `bool` | `True` | 是否匹配有大小的元素，`False`表示匹配没有大小的 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 匹配成功返回元素对象 |
| `NoneElement` | 失败返回`NoneElement` |
| 抛出`ElementNotFoundError`异常 | `Settings.raise_when_ele_not_found`为`True`时抛出 |

---

### 📌 `filter_one.style()`[​](#-filter_onestyle "-filter_onestyle的直接链接")

此方法用于在元素列表中以是否拥有某个 style 值为条件筛选元素，返回第一个结果。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `name` | `str` | 必填 | 属性名称 |
| `value` | `str` | 必填 | 属性值 |
| `equal` | `bool` | `True` | `True`表示匹配`name`值为`value`值的元素，`False`表示匹配`name`值不为`value`值的 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 匹配成功返回元素对象 |
| `NoneElement` | 失败返回`NoneElement` |
| 抛出`ElementNotFoundError`异常 | `Settings.raise_when_ele_not_found`为`True`时抛出 |

---

### 📌 `filter_one.property()`[​](#-filter_oneproperty "-filter_oneproperty的直接链接")

此方法用于在元素列表中以是否拥有某个 property 值为条件筛选元素，返回第一个结果。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `name` | `str` | 必填 | 属性名称 |
| `value` | `str` | 必填 | 属性值 |
| `equal` | `bool` | `True` | `True`表示匹配`name`值为`value`值的元素，`False`表示匹配`name`值不为`value`值的 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 匹配成功返回元素对象 |
| `NoneElement` | 失败返回`NoneElement` |
| 抛出`ElementNotFoundError`异常 | `Settings.raise_when_ele_not_found`为`True`时抛出 |

---

### 📌 `filter_one.attr()`[​](#-filter_oneattr "-filter_oneattr的直接链接")

此方法用于在元素列表中以是否拥有某个 attribute 值为条件筛选元素，返回第一个结果。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `name` | `str` | 必填 | 属性名称 |
| `value` | `str` | 必填 | 属性值 |
| `equal` | `bool` | `True` | `True`表示匹配`name`值为`value`值的元素，`False`表示匹配`name`值不为`value`值的 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 匹配成功返回元素对象 |
| `NoneElement` | 失败返回`NoneElement` |
| 抛出`ElementNotFoundError`异常 | `Settings.raise_when_ele_not_found`为`True`时抛出 |

---

### 📌 `filter_one.text()`[​](#-filter_onetext "-filter_onetext的直接链接")

此方法用于在元素列表中以是否含有指定文本为条件筛选元素，返回第一个结果。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `text` | `str` | 必填 | 用于匹配的文本 |
| `fuzzy` | `bool` | `True` | 是否模糊匹配 |
| `contain` | `bool` | `True` | 是否包含该字符串，`False`表示不包含 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 匹配成功返回元素对象 |
| `NoneElement` | 失败返回`NoneElement` |
| 抛出`ElementNotFoundError`异常 | `Settings.raise_when_ele_not_found`为`True`时抛出 |

---

### 📌 `filter_one.tag()`[​](#-filter_onetag "-filter_onetag的直接链接")

此方法用于在元素列表中以某个类型为条件筛选元素，返回第一个结果。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `name` | `str` | 必填 | 元素类型名称 |
| `equal` | `bool` | `True` | `True`表示匹配该类型元素，`False`表示匹配非该类型元素 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 匹配成功返回元素对象 |
| `NoneElement` | 失败返回`NoneElement` |
| 抛出`ElementNotFoundError`异常 | `Settings.raise_when_ele_not_found`为`True`时抛出 |

---

### 📌 选择获取第几个[​](#-选择获取第几个 "📌 选择获取第几个的直接链接")

`filter_one`可加参数，以选择返回第几个结果。

**示例：**

```
ele = eles.filter_one(2).text('图')  # 获取第二个文本带有“图”字的元素
```

说明

`filter_one`在不加序号参数时，可不要后面的`()`。

---

✅️️ 获取全部匹配元素[​](#️️-获取全部匹配元素 "✅️️ 获取全部匹配元素的直接链接")
-------------------------------------------------

说明

静态元素列表只有`filter.attr()`和`filter.text()`方法。

### 📌 `filter.displayed()`[​](#-filterdisplayed "-filterdisplayed的直接链接")

此方法用于在元素列表中以是否显示为条件筛选元素，返回新的列表。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `equal` | `bool` | `True` | 是否匹配显示的元素，`False`匹配不显示的 |

| 返回类型 | 说明 |
| --- | --- |
| `Filter` | 元素对象组成的列表，可继续用于筛选 |

---

### 📌 `filter.checked()`[​](#-filterchecked "-filterchecked的直接链接")

此方法用于在元素列表中以是否被选中为条件筛选元素，返回新的列表。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `equal` | `bool` | `True` | 是否匹配被选中的元素，`False`匹配不被选中的 |

| 返回类型 | 说明 |
| --- | --- |
| `Filter` | 元素对象组成的列表，可继续用于筛选 |

---

### 📌 `filter.selected()`[​](#-filterselected "-filterselected的直接链接")

此方法用于在元素列表中以是否被选择为条件筛选元素，返回新的列表。用于`<select>`元素项目。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `equal` | `bool` | `True` | 是否匹配被选择的元素，`False`匹配不被选择的 |

| 返回类型 | 说明 |
| --- | --- |
| `Filter` | 元素对象组成的列表，可继 续用于筛选 |

---

### 📌 `filter.enabled()`[​](#-filterenabled "-filterenabled的直接链接")

此方法用于在元素列表中以是否可用为条件筛选元素，返回新的列表。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `equal` | `bool` | `True` | 是否匹配可用的元素，`False`匹配 disabled 状态的 |

| 返回类型 | 说明 |
| --- | --- |
| `Filter` | 元素对象组成的列表，可继续用于筛选 |

---

### 📌 `filter.clickable()`[​](#-filterclickable "-filterclickable的直接链接")

此方法用于在元素列表中以是否可点击为条件筛选元素，返回新的列表。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `equal` | `bool` | `True` | 是否匹配可点击的元素，`False`表示匹配不是可点击的 |

| 返回类型 | 说明 |
| --- | --- |
| `Filter` | 元素对象组成的列表，可继续用于筛选 |

---

### 📌 `filter.have_rect()`[​](#-filterhave_rect "-filterhave_rect的直接链接")

此方法用于在元素列表中以是否有大小为条件筛选元素，返回新的列表。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `equal` | `bool` | `True` | 是否匹配有大小的元素，`False`表示匹配没有大小的 |

| 返回类型 | 说明 |
| --- | --- |
| `Filter` | 元素对象组成的列表，可继续用于筛选 |

---

### 📌 `filter.style()`[​](#-filterstyle "-filterstyle的直接链接")

此方法用于在元素列表中以是否拥有某个 style 值为条件筛选元素，返回新的列表。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `name` | `str` | 必填 | 属性名称 |
| `value` | `str` | 必填 | 属性值 |
| `equal` | `bool` | `True` | `True`表示匹配`name`值为`value`值的元素，`False`表示匹配`name`值不为`value`值的 |

| 返回类型 | 说明 |
| --- | --- |
| `Filter` | 元素对象组成的列表，可继续用于筛选 |

---

### 📌 `filter.property()`[​](#-filterproperty "-filterproperty的直接链接")

此方法用于在元素列表中以是否拥有某个 property 值为条件筛选元素，返回新的列表。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `name` | `str` | 必填 | 属性名称 |
| `value` | `str` | 必填 | 属性值 |
| `equal` | `bool` | `True` | `True`表示匹配`name`值为`value`值的元素，`False`表示匹配`name`值不为`value`值的 |

| 返回类型 | 说明 |
| --- | --- |
| `Filter` | 元素对象组成的列表，可继续用于筛选 |

---

### 📌 `filter.tag()`[​](#-filtertag "-filtertag的直接链接")

此方法用于在元素列表中以某个类型为条件筛选元素，返回新的列表。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `name` | `str` | 必填 | 元素类型名称 |
| `equal` | `bool` | `True` | `True`表示匹配该类型元素，`False`表示匹配非该类型元素 |

| 返回类型 | 说明 |
| --- | --- |
| `Filter` | 元素对象组成的列表，可继续用于筛选 |

---

### 📌 `filter.attr()`[​](#-filterattr "-filterattr的直接链接")

此方法用于在元素列表中以是否拥有某个 attribute 值为条件筛选元素，返回新的列表。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `name` | `str` | 必填 | 属性名称 |
| `value` | `str` | 必填 | 属性值 |
| `equal` | `bool` | `True` | `True`表示匹配`name`值为`value`值的元素，`False`表示匹配`name`值不为`value`值的 |

| 返回类型 | 说明 |
| --- | --- |
| `Filter` | 元素对象组成的列表，可继续用于筛选 |

---

### 📌 `filter.text()`[​](#-filtertext "-filtertext的直接链接")

此方法用于在元素列表中以是否含有指定文本为条件筛选元素，返回新的列表。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `text` | `str` | 必填 | 用于匹配的文本 |
| `fuzzy` | `bool` | `True` | 是否模糊匹配 |
| `contain` | `bool` | `True` | 是否包含该字符串，`False`表示不包含 |

| 返回类型 | 说明 |
| --- | --- |
| `Filter` | 元素对象组成的列表，可继续用于筛选 |

---

✅️️ 多条件筛选[​](#️️-多条件筛选 "✅️️ 多条件筛选的直接链接")
----------------------------------------

### 📌 与关系筛选[​](#-与关系筛选 "📌 与关系筛选的直接链接")

筛选支持链式操作，可串连多个条件，每个条件会筛选一层再进入下一层。

可实现多个条件的与关系筛选。

**示例，出导航栏中显示且含有“图”字的元素：**

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.get('https://www.baidu.com')  
eles = tab('#s-top-left').eles('t:a')  
for ele in eles.filter.displayed().text('图'):  
    print(ele.text, end=' ')
```

---

### 📌 或关系筛选[​](#-或关系筛选 "📌 或关系筛选的直接链接")

元素列表的`search()`和`search_one()`方法可用于多个条件或筛选元素。

说明

静态元素列表没有这种方法。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `index`（`search_one()`独有） | `int` | `1` | 结果中的元素序号，`1`开始 |
| `displayed` | `bool` | `None` | 是否显示，`bool`表示匹配是或否，`None`为忽略该项 |
| `checked` | `bool` | `None` | 是否被选中，`bool`表示匹配是或否，`None`为忽略该项 |
| `selected` | `bool` | `None` | 是否被选择，`bool`表示匹配是或否，`None`为忽略该项 |
| `enabled` | `bool` | `None` | 是否可用，`bool`表示匹配是或否，`None`为忽略该项 |
| `clickable` | `bool` | `None` | 是否可点击，`bool`表示匹配是或否，`None`为忽略该项 |
| `have_rect` | `bool` | `None` | 是否拥有大小，`bool`表示匹配是或否，`None`为忽略该项 |
| `have_text` | `bool` | `None` | 是否含有文本，`bool`表示匹配是或否，`None`为忽略该项 |
| `tag` | `str` | `None` | 指定标签页类型，`None`为忽略该项 |

| 返回类型 | 说明 |
| --- | --- |
| `Filter` | `search()`返回元素对象组成的列表 |
| `ChromiumElement` | `search_one()`匹配成功返回元素对象 |
| `NoneElement` | `search_one()`匹配失败返回`NoneElement` |

---

### 📌 混合筛选[​](#-混合筛选 "📌 混合筛选的直接链接")

与关系和或关系可以用链式操作混合使用。

说明

静态元素列表没有这种方法。

```
es = eles.search(displayed=True).enabled()  
ele = eles.filter.enablde().search_one(displayed=True)
```

---

✅️️ 获取筛选结果的属性集合[​](#️️-获取筛选结果的属性集合 "✅️️ 获取筛选结果的属性集合的直接链接")
----------------------------------------------------------

筛选结果列表可以调用`get()`方法获取指定属性结合。

该集合为遍历列表中所有元素获取的。

**示例：**

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.get('https://www.baidu.com')  
eles = tab('#s-top-left').eles('t:a')  
print(eles.get.texts())  # 获取所有元素的文本  
print(eles.filter.displayed().get.texts())  # 获取的元素的文本
```

**输出：**

```
['新闻', 'hao123', '地图', '贴吧', '视频', '图片', '网盘', '文库', '更多', '翻译', '学术', '百科', '知道', '健康', '营销推广', '直播', '音乐', '橙篇', '查看全部百度产品 >']  
['新闻', 'hao123', '地图', '贴吧', '视频', '图片', '网盘', '文库', '更多']
```

### 📌 `get.attrs()`[​](#-getattrs "-getattrs的直接链接")

此方法用于返回所有元素指定的 attribute 属性组成的列表。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `name` | `str` | 必填 | 属性名称 |

| 返回类型 | 说明 |
| --- | --- |
| `List[str]` | 属性文本组成的列表 |

---

### 📌 `get.links()`[​](#-getlinks "-getlinks的直接链接")

此方法用于返回所有元素的`link`属性组成的列表。

**参数：** 无

| 返回类型 | 说明 |
| --- | --- |
| `List[str]` | 链接文本组成的列表 |

---

### 📌 `get.texts()`[​](#-gettexts "-gettexts的直接链接")

此方法用于返回所有 元素的`text`属性组成的列表。

**参数：** 无

| 返回类型 | 说明 |
| --- | --- |
| `List[str]` | 元素文本组成的列表 |

[上一页

🔦 行为模式](/browser_control/get_elements/behavior)[下一页

🔦 简化写法](/browser_control/get_elements/simplify)

* [✅️️ 获取单个匹配元素](#️️-获取单个匹配元素)
  + [📌 `filter_one.displayed()`](#-filter_onedisplayed)
  + [📌 `filter_one.checked()`](#-filter_onechecked)
  + [📌 `filter_one.selected()`](#-filter_oneselected)
  + [📌 `filter_one.enabled()`](#-filter_oneenabled)
  + [📌 `filter_one.clickable()`](#-filter_oneclickable)
  + [📌 `filter_one.have_rect()`](#-filter_onehave_rect)
  + [📌 `filter_one.style()`](#-filter_onestyle)
  + [📌 `filter_one.property()`](#-filter_oneproperty)
  + [📌 `filter_one.attr()`](#-filter_oneattr)
  + [📌 `filter_one.text()`](#-filter_onetext)
  + [📌 `filter_one.tag()`](#-filter_onetag)
  + [📌 选择获取第几个](#-选择获取第几个)
* [✅️️ 获取全部匹配元素](#️️-获取全部匹配元素)
  + [📌 `filter.displayed()`](#-filterdisplayed)
  + [📌 `filter.checked()`](#-filterchecked)
  + [📌 `filter.selected()`](#-filterselected)
  + [📌 `filter.enabled()`](#-filterenabled)
  + [📌 `filter.clickable()`](#-filterclickable)
  + [📌 `filter.have_rect()`](#-filterhave_rect)
  + [📌 `filter.style()`](#-filterstyle)
  + [📌 `filter.property()`](#-filterproperty)
  + [📌 `filter.tag()`](#-filtertag)
  + [📌 `filter.attr()`](#-filterattr)
  + [📌 `filter.text()`](#-filtertext)
* [✅️️ 多条件筛选](#️️-多条件筛选)
  + [📌 与关系筛选](#-与关系筛选)
  + [📌 或关系筛选](#-或关系筛选)
  + [📌 混合筛选](#-混合筛选)
* [✅️️ 获取筛选结果的属性集合](#️️-获取筛选结果的属性集合)
  + [📌 `get.attrs()`](#-getattrs)
  + [📌 `get.links()`](#-getlinks)
  + [📌 `get.texts()`](#-gettexts)

* 🚀 控制浏览器
* 🔎 查找元素
* 🔦 页面或元素内查找

本页总览

🔦 页面或元素内查找
==========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

✅️️ 页面或元素内查找[​](#️️-页面或元素内查找 "✅️️ 页面或元素内查找的直接链接")
-------------------------------------------------

页面对象和元素对象都拥有`ele()`和`eles()`方法，用于获取其内部指定子元素。

### 📌 `ele()`[​](#-ele "-ele的直接链接")

用于查找其内部第一个符合条件的元素。

`SessionPage`和`ChromiumPage`获取元素的方法是一致的，但前者返回的元素对象为`SessionElement`，后者是`ChromiumElement`。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` | 必填 | 元素的定位信息。可以是查询字符串，或 loc 元组 |
| `index` | `int` | `1` | 获取第几个匹配的元素，从`1`开始，可输入负数表示从后面开始数 |
| `timeout` | `float` | `None` | 等待元素出现的超时时间（秒），为`None`使用页面对象设置 |

| 返回类型 | 说明 |
| --- | --- |
| `SessionElement` | s 模式下返回静态元素对象 |
| `ChromiumElement` | d 模式下返回找到的第一个符合条件的浏览器元素对象 |
| `ChromiumFrame` | 当结果是框架元素时 |
| `NoneElement` | 未找到符合条件的元素时返回 |

说明

* loc 元组是指 selenium 定位符，例：(By.ID, '\*\*\*\*')。下同。
* `ele('****', index=2)`和`eles('****')[1]`结果一样，不过前者会快很多。

**示例：**

```
from DrissionPage import SessionPage  
  
page = SessionPage()  
  
# 在页面内查找元素  
ele1 = page.ele('#one')  
  
# 在元素内查找后代元素  
ele2 = ele1.ele('第二行')
```

---

### 📌 `eles()`[​](#-eles "-eles的直接链接")

此方法与`ele()`相似，但返回的是匹配到的所有元素组成的列表。

页面对象和元素对象都可调用这个方法。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` | 必填 | 元素的定位信息，可以是查询字符串，或 loc 元组 |
| `timeout` | `float` | `None` | 等待元素出现的超时时间（秒），为`None`使用页面对象设置 |

| 返回类型 | 说明 |
| --- | --- |
| `SessionElementsList` | s 模式下返回静态元素对象组成的列表 |
| `ChromiumElementsList` | d 模式下返回浏览器元素对象组成的列表 |

**示例：**

```
# 获取页面内的所有p元素  
p_eles = tab.eles('tag:p')  
  
# 获取ele1元素内的所有p元素  
p_eles = ele1.eles('tag:p')  
  
# 打印第一个p元素的文本  
print(p_eles[0])
```

---

### 📌 `get_frame()`[​](#-get_frame "-get_frame的直接链接")

`<iframe>`和`<frame>`也可以用`ele()`查找到，生成的对象是`ChromiumFrame`而不是`ChromiumElement`。

但不建议用`ele()`获取`<iframe>`元素，因为 IDE 无法正确提示后续操作。

建议用`get_frame()`方法获取，页面对象和元素对象都有这个方法。

使用方法与`ele()`一致，可以用定位符查找。还增加了用序号、id、name 属性定位元素的功能。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `loc_ind_ele` | `str` `int` `ChromiumFrame` | 必填 | 定位符 `<iframe>`元素序号（从`1`开始，负数表示倒数） `ChromiumFrame对象` `id`属性内容 `name`属性内容 |
| `timeout` | `float` | `None` | 超时时间（秒），为`None`时使用页面超时时间 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumFrame` | `<frame>`或`<iframe>`元素对象 |
| `NoneElement` | 找不到时返回`NoneElement` |

**示例：**

```
# 在标签页中获取第一个iframe元素  
iframe = tab.get_frame(1)  
  
# 在元素中获取id为`theFrame`的<iframe>元素对象  
iframe = ele.get_frame('#theFrame')
```

---

### 📌 `get_frames()`[​](#-get_frames "-get_frames的直接链接")

此方法用于获取页面中多个符合条件的`<frame>`或`<iframe>`对象。

元素对象无此方法。

提醒

获取所有`<iframe>`会很慢，而且浪费资源，非必要别用。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` | `None` | 定位符，为`None`时返回所有 |
| `timeout` | `float` | `None` | 超时时间（秒），为`None`时使用页面超时时间 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElementsList` | `<frame>`或`<iframe>`元素对象组成的列表 |

---

✅️️ 静态方式查找[​](#️️-静态方式查找 "✅️️ 静态方式查找的直接链接")
-------------------------------------------

静态元素即 s 模式的`SessionElement`元素对象，是纯文本构造的，因此用它处理速度非常快。  
对于复杂的页面，要在成百上千个元素中采集数据时，转换为静态元素可把速度提升几个数量级。  
作者曾在实践的时候，用同一套逻辑，仅仅把元素转换为静态，就把一个要 30 秒才完成的页面，加速到零点几秒完成。  
我们甚至可以把整个页面转换为静态元素，再在其中提取信息。  
当然，这种元素不能进行点击等交互。  
用`s_ele()`可在把查找到的动态元素转换为静态元素输出，或者获取元素或页面本身的静态元素副本。

注意

如果需要获取多条数据，不要反复使用`s_ele()`，只要在容器元素调用一次获取其静态副本，再在其中执行获取多个元素。

### 📌 `s_ele()`[​](#-s_ele "-s_ele的直接链接")

页面对象和元素对象都拥有此方法，用于查找第一个匹配条件的元素，获取其静态版本。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` | `None` | 元素的定位信息，可以是查询字符串，或 loc 元组，为`None`时获取调用者本身的静态版本 |
| `index` | `int` | `1` | 获取第几个匹配的元素，从`1`开始，可输入负数表示从后面开始数 |
| `timeout` | `float` | `None` | 查找元素超时时间（秒），为`None`与页面等待时间一致 |

| 返回类型 | 说明 |
| --- | --- |
| `SessionElement` | 返回查找到的第一个符合条件的元素对象的静态版本 |
| `NoneElement` | 限时内未找到符合条件的元素时返回`NoneElement`对象 |

注意

页面对象和元素对象的`s_ele()`方法不能搜索到在`<iframe>`里的元素，页面对象的静态版本也不能搜索`<iframe>`里的元素。
要使用`<iframe>`里元素的静态版本，可先获取该元素，再转换。而使用`ChromiumFrame`对象，则可以直接用`s_ele()`查找元素，这在后面章节再讲述。

Tips

从一个`ChromiumElement`元素获取到的`SessionElement`版本，依然能够使用相对定位方法定位祖先或兄弟元素。

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
  
# 在页面中查找元素，获取其静态版本  
ele1 = tab.s_ele('search text')  
  
# 在动态元素中查找元素，获取其静态版本  
ele = tab.ele('search text')  
ele2 = ele.s_ele()  
  
# 获取页面元素的静态副本（不传入参数）  
s_page = tab.s_ele()  
  
# 获取动态元素的静态副本  
s_ele = ele.s_ele()  
  
# 在静态副本中查询下级元素（因为已经是静态元素，用ele()查找结果也是静态）  
ele3 = s_page.ele('search text')  
ele4 = s_ele.ele('search text')
```

---

### 📌 `s_eles()`[​](#-s_eles "-s_eles的直接链接")

此方法与`s_ele()`相似，但返回的是匹配到的所有元素组成的列表，或属性值组成的列表。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` | 必填 | 元素的定位信息，可以是查询字符串，或 loc 元组 |
| `timeout` | `float` | `None` | 查找元素超时时间（秒），为`None`与页面等待时间一致 |

| 返回类型 | 说明 |
| --- | --- |
| `SessionElementsList` | 返回找到的所有元素的`SessionElement`版本组成的列表 |

**示例：**

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
for ele in tab.s_eles('search text'):  
    print(ele.text)
```

---

✅️ 获取页面焦点元素[​](#️-获取页面焦点元素 "✅️ 获取页面焦点元素的直接链接")
----------------------------------------------

使用页面对象的`active_ele`属性获取页面上焦点所在元素。

```
ele = tab.active_ele
```

---

  

✅️️ 查找`<iframe>`中的元素[​](#️️-查找iframe中的元素 "️️-查找iframe中的元素的直接链接")
----------------------------------------------------------------

### 📌 在页面下跨级查找[​](#-在页面下跨级查找 "📌 在页面下跨级查找的直接链接")

与 selenium 不同，本库可以直接查找同域`<iframe>`里面的元素。  
而且无视层级，可以直接获取到多层`<iframe>`里的元素。无需切入切出，大大简化了程序逻辑，使用更便捷。  
但同域的`<iframe>`才能这样查找。

假设在页面中有个两级`<iframe>`，其中有个元素`<div id='abc'></div>`，可以这样获取：

```
tab = Chromium().latest_tab  
ele = tab('#abc')
```

获取前后无需切入切出，也不影响获取页面上其它元素。

如果用 selenium，要这样写：

```
driver = webdriver.Chrome()  
driver.switch_to.frame(0)  
driver.switch_to.frame(0)  
ele = driver.find_element(By.ID, 'abc')  
driver.switch_to.default_content()
```

显然比较繁琐，而且切入到`<iframe>`后无法对`<iframe>`外的元素进行操作。

注意

* 跨级查找只是页面对象支持，元素对象不能直接查找内部 iframe 里的元素。
* 跨级查找只能用于与主框架同域 名的`<iframe>`，不同域名的请用下面的方法。

---

### 📌 在 iframe 元素内查找[​](#-在-iframe-元素内查找 "📌 在 iframe 元素内查找的直接链接")

对于跨域的`<iframe>`，我们无法通过页面直接查找里面的元素，可以先获取到`<iframe>`元素，再在其下查找。

当然，非跨域`<iframe>`也可以这样操作。

假设一个`<iframe>`的 id 为 `'iframe1'`，要在其中查找一个 id 为`'abc'`的元素：

```
tab = Chromium().latest_tab  
iframe = tab('#iframe1')  
ele = iframe('#abc')
```

这个`<iframe>`元素是一个页面对象，因此可以继续在其下进行跨`<iframe>`查找（相对这个`<iframe>`不跨域的）。

---

✅️️ `ShadowRoot`内查找[​](#️️-shadowroot内查找 "️️-shadowroot内查找的直接链接")
-----------------------------------------------------------------

本库把 shadow-root 也作为元素对象看待，是为`ShadowRoot`对象。
该对象可与普通元素一样查找下级元素和 DOM 内相对定位。  
对`ShadowRoot`对象进行相对定位时，把它看作其父对象内部的第一个对象，其余定位逻辑与普通对象一致。

用元素对象的`shadow_root`属性可获取`ShadowRoot`对象。

注意

如果`ShadowRoot`元素的下级元素中有其它`ShadowRoot`元素，那这些下级`ShadowRoot`
元素内部是无法直接通过定位语句查找到的，只能先定位到其父元素，再用`shadow-root`属性获取。

```
# 获取一个 shadow-root 元素  
sr_ele = page.ele('#app').shadow_root  
  
# 在该元素下查找下级元素  
ele1 = sr_ele.ele('tag:div')  
  
# 用相对定位获取其它元素  
ele1 = sr_ele.parent(2)  
ele1 = sr_ele.next('tag:div', 1)  
ele1 = sr_ele.after('tag:div', 1)  
eles = sr_ele.nexts('tag:div')  
  
# 定位下级元素中的 shadow+-root 元素  
sr_ele2 = sr_ele.ele('tag:div').shadow_root
```

由于 shadow-root 不能跨级查找，链式操作非常常见，所以设计了一个简写：`sr`，功能和`shadow_root`
一样，都是获取元素内部的`ShadowRoot`。

**多级 shadow-root 链式操作示例：**

以下这段代码，可以打印浏览器历史第一页，可见是通过多级 shadow-root 来获取的。

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.get('chrome://history/')  
  
items = tab('#history-app').sr('#history').sr.eles('t:history-item')  
for i in items:  
    print(i.sr('#item-container').text.replace('\n', ''))
```

---

✅️️ 同时匹配多个定位符[​](#️️-同时匹配多个定位符 "✅️️ 同时匹配多个定位符的直接链接")
----------------------------------------------------

所有页面或元素对象都有`find()`方法，可接收多个定位符，同时查找多个（批）不同定位符的元素。

以`dict`方法返回每个定位符结果。

说明

当`first_ele`为`True`时，如果一个定位符没有被执行过查找，它返回的结果为`None`。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locators` | `List[str]` `Tuple[str, str]` `str` | 必填 | 一个定位符或多个组成的列表 |
| `any_one` | `bool` | `True` | 是否任何一个定位符找到结果即返回 |
| `first_ele` | `bool` | `True` | 每个定位符获取第一个元素还是所有元素 |
| `timeout` | `float` | `None` | 超时时间（秒），为`None`使用该对象默认设置 |

说明

以  下所说的 “定位符”，是`str`或`tuple`类型的。
“元素对象”，是`ChromiumElement`（d 模式）或`SessionElement`（s 模式）类型的，没有找到时是`NoneElement`类型的。
“元素对象组成的列表” 是`ChromiumElementsList`（d 模式）或`SessionElementsList`（s 模式）类型的。
`any_one`参数为`True`时，以`tuple`方式返回找到目标的定位符和结果，为`False`时以`dict`方法返回每个定位符结果。

| 返回类型 | `any_one`参数取值 | 说明 |
| --- | --- | --- |
| `tuple(定位符, 元素对象)` | `True` | `first_ele`为`True`时，返回第一个有结果的定位符找到的第一个元素对象 |
| `tuple(定位符, 元素对象组成的列表)` | `True` | `first_ele`为`False`时，返回第一个有结果的定位符找到的所有元素对象 |
| `tuple(None, None)` | `True` | 所有定位符都没有找到元素，返回`(None, None)` |
| `dict{定位符: 元素对象}` | `False` | `first_ele`为`True`时，每个定位符返回第一个元素，找不到时为`NoneElement` |
| `dict{定位符: 元素对象组成的列表}` | `False` | `first_ele`为`False`时，每个定位符返回所有结果  元素组成的列表 |

**示例：**

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.get('https://www.baidu.com')  
res = tab.find(['#kw', '#su'])  
print(res)
```

[上一页

🔦 定位语法](/browser_control/get_elements/syntax)[下一页

🔦 相对定位](/browser_control/get_elements/relative)

* [✅️️ 页面或元素内查找](#️️-页面或元素内查找)
  + [📌 `ele()`](#-ele)
  + [📌 `eles()`](#-eles)
  + [📌 `get_frame()`](#-get_frame)
  + [📌 `get_frames()`](#-get_frames)
* [✅️️ 静态方式查找](#️️-静态方式查找)
  + [📌 `s_ele()`](#-s_ele)
  + [📌 `s_eles()`](#-s_eles)
* [✅️ 获取页面焦点元素](#️-获取页面焦点元素)
* [✅️️ 查找`<iframe>`中的元素](#️️-查找iframe中的元素)
  + [📌 在页面下跨级查找](#-在页面下跨级查找)
  + [📌 在 iframe 元素内查找](#-在-iframe-元素内查找)
* [✅️️ `ShadowRoot`内查找](#️️-shadowroot内查找)
* [✅️️ 同时匹配多个定位符](#️️-同时匹配多个定位符)

* 🚀 控制浏览器
* 🔎 查找元素
* 🔦 概述

本页总览

🔦 概述
====

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

定位元素是自动化重中之重的技能。  
虽然可在开发者工具直接复制绝对路径，但这样做有几个缺点：

* 代码冗长，可读性低
* 动态页面容易导致元素失效
* 无法使用相对定位
* 网页稍有改动或者出现临时元素就不能用，容错性低
* 无法跨`<iframe>`查找元素

因此作者极不建议使用右键复制的元素路径。

本库提供一套简洁易用的语法，用于快速定位元素，并且内置等待功能、支持链式查找，减少了代码的复杂性。  
同时也兼容 css selector、xpath 和 selenium 原生的 loc 元组。

✅️️ 基本用法[​](#️️-基本用法 "✅️️ 基本用法的直接链接")
-------------------------------------

所有页面对象和元素对象（包括`<iframe>`和 shadow-root），都可以在自己内部查找元素。

元素对象还能以自己为基准，相对定位其它元素。

定位元素大致有以下几种方法，将在后续小节中详细说明。

* 在页面或元素内查找子元素
* 根据 DOM 结构相对定位
* 根据视觉位置相对定位

所有的查找元素方法，都可以使用本库自创的查找语法、xpath、css selector和 selenium 的定位符元组，去查找元素。

---

✅️️ 示例[​](#️️-示例 "✅️️ 示例的直接链接")
-------------------------------

### 📌 简单示例[​](#-简单示例 "📌 简单示例的直接链接")

假设有这样一个页面：

```
<html>  
<body>  
<div id="one">  
    <p class="p_cls" name="row1">第一行</p>  
    <p class="p_cls" name="row2">第二行</p>  
    <p class="p_cls">第三行</p>  
</div>  
<div id="two">  
    第二个div  
</div>  
</body>  
</html>
```

我们可以用页面对象去获取其中的元素：

```
div1 = tab.ele('#one')  # 获取 id 为 one 的元素  
p1 = tab.ele('@name=row1')  # 获取 name 属性为 row1 的元素  
div2 = tab.ele('第二个div')  # 获取包含“第二个div”文本的元素  
div_list = tab.eles('tag:div')  # 获取所有div元素
```

也可以获取到一个元素，然后在它里面或周围查找元素：

```
div1 = tab.ele('#one')  # 获取到一个元素div1  
p_list = div1.eles('tag:p')  # 在div1内查找所有p元素  
div2 = div1.next()  # 获取div1后面一个元素
```

---

### 📌 实际示例[​](#-实际示例 "📌 实际示例的直接链接")

复制此代码可直接运行查看结果。

```
from DrissionPage import SessionPage  
  
page = SessionPage()  
page.get('https://gitee.com/explore')  
  
# 获取包含“全部推荐项目”文本的 ul 元素  
ul_ele = page.ele('tag:ul@text():全部推荐项目')    
  
# 获取该 ul 元素下所有 a 元素  
titles = ul_ele.eles('tag:a')    
  
# 遍历列表，打印每个 a 元素的文本  
for i in titles:    
    print(i.text)
```

**输出：**

```
全部推荐项目  
前沿技术  
智能硬件  
IOT/物联网/边缘计算  
车载应用  
...
```

[上一页

🛰️ 获取网页信息](/browser_control/get_page_info)[下一页

🔦 定位语法](/browser_control/get_elements/syntax)

* [✅️️ 基本用法](#️️-基本用法)
* [✅️️ 示例](#️️-示例)
  + [📌 简单示例](#-简单示例)
  + [📌 实际示例](#-实际示例)

* 🚀 控制浏览器
* 🔎 查找元素
* 🔦 相对定位

本页总览

🔦 相对定位
======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

相对定位的意思是以一个已获取的元素为基准，按需要使用不同方法获取指定的其它元素。

相对定位有基于 DOM 的方式和基于视觉的方式两种。

✅️️ 基于 DOM 相对定位[​](#️️-基于-dom-相对定位 "✅️️ 基于 DOM 相对定位的直接链接")
----------------------------------------------------------

以下方法可以以某元素为基准，在 DOM 中按照条件获取其直接子节点、同级节点、祖先元素、文档前后节点。

这里说的是 “节点”，不是 “元素”。因为相对定位可以获取除元素外的其它节点，包括文本、注释节点。

注意

如果元素在`<iframe>`中，相对定位不能超越`<iframe>`文档。

### 📌 获取父级元素[​](#-获取父级元素 "📌 获取父级元素的直接链接")

🔸 `parent()`

此方法获取当前元素某一级父元素，可指定筛选条件或层数。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `level_or_loc` | `int` `str` `Tuple[str, str]` | `1` | 第几级父元素，从`1`开始，或用定位符在祖先元素中进行筛选 |
| `index` | `int` | `1` | 当`level_or_loc`传入定位符，使用此参数选择第几个结果，从当前元素往上级数；当`level_or_loc`传入数字时，此参数无效 |
| `timeout` | `float` | `0` | 查找超时时间（秒），为`None`时使用页面超时设置，s 模式下无效 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | d 模式下的元素对象 |
| `SessionElement` | s 模式下的元素对象 |
| `NoneElement` | 未获取到结果时 |

**示例：**

```
# 获取 ele1 的第二层父元素  
ele2 = ele1.parent(2)  
  
# 获取 ele1 父元素中 id 为 id1 的元素  
ele2 = ele1.parent('#id1')
```

---

### 📌 获取直接子节点[​](#-获取直接子节点 "📌 获取直接子节点的直接链接")

🔸 `child()`

此方法返回当前元素的一个直接子节点，可指定筛选条件和第几个。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` `int` | `''` | 用于筛选节点的查询语法，为`int`类型时`index`参数无效 |
| `index` | `int` | `1` | 查询结果中的第几个，从`1`开始，可输入负数表示倒数 |
| `timeout` | `float` | `None` | 查找超时时间（秒），为`None`时使用页面超时设置，s 模式下无效 |
| `ele_only` | `bool` | `True` | 是否只查找元素，为`False`时把文本、注释节点也纳入查找范围 |

| 返回类型 | 说明 |
| --- | --- |
| `str` | 获取非元素节点时返回字符串 |
| `ChromiumElement` | d 模式下的元素对象 |
| `SessionElement` | s 模式下的元素对象 |
| `NoneElement` | 未获取到结果时 |

---

🔸 `children()`

此方法返回当前元素全部符合条件的直接子节点组成的列表，可用查询语法筛选。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` | `''` | 用于筛选节点的查询语法 |
| `timeout` | `float` | `None` | 查找超时时间（秒），为`None`时使用页面超时设置，s 模式下无效 |
| `ele_only` | `bool` | `True` | 是否只查找元素，为`False`时把文本、注释节点也纳入查找范围 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElementsList` | d 模式结果列表 |
| `SessionElementsList` | s 模式结果列表 |

---

### 📌 获取后面的同级节点[​](#-获取后面的同级节点 "📌 获取后面的同级节点的直接链接")

🔸 `next()`

此方法返回当前元素后面的某一个同级节点，可指定筛选条件和第几个。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` `int` | `''` | 用于筛选节点的查询语法，为`int`类型时`index`参数无效 |
| `index` | `int` | `1` | 查询结果中的第几个，从`1`开始，可输入负数表示倒数 |
| `timeout` | `float` | `None` | 查找超时时间（秒），为`None`时使用页面超时设置，s 模式下无效 |
| `ele_only` | `bool` | `True` | 是否只查找元素，为`False`时把文本、注释节点也纳入查找范围 |

| 返回类型 | 说明 |
| --- | --- |
| `str` | 获取非元素节点时返回字符串 |
| `ChromiumElement` | d 模式下的元素对象 |
| `SessionElement` | s 模式下的元素对象 |
| `NoneElement` | 未获取到结果时 |

**示例：**

```
# 获取 ele1 后面第一个兄弟元素  
ele2 = ele1.next()  
  
# 获取 ele1 后面第 3 个兄弟元素  
ele2 = ele1.next(3)  
  
# 获取 ele1 后面第 3 个 div 兄弟元素  
ele2 = ele1.next('tag:div', 3)  
  
# 获取 ele1 后面第一个文本节点的文本  
txt = ele1.next('xpath:text()', 1)
```

---

🔸 `nexts()`

此方法返回当前元素后面全部符合条件的同级节点组成的列表，可用查询语法筛选。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` | `''` | 用于筛选节点的查询语法 |
| `timeout` | `float` | `None` | 查找超时时间（秒），为`None`时使用页面超时设置，s 模式下无效 |
| `ele_only` | `bool` | `True` | 是否只查找元素，为`False`时把文本、注释节点也纳入查找范围 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElementsList` | d 模式结果列表 |
| `SessionElementsList` | s 模式结果列表 |

**示例：**

```
# 获取 ele1 后面所有兄弟元素  
eles = ele1.nexts()  
  
# 获取 ele1 后面所有 div 兄弟元素  
divs = ele1.nexts('tag:div')  
  
# 获取 ele1 后面的所有文本节点  
txts = ele1.nexts('xpath:text()')
```

---

### 📌 获取前面的同级节点[​](#-获取前面的同级节点 "📌 获取前面的同级节点的直接链接")

🔸 `prev()`

此方法返回当前元素前面的某一个同级节点，可指定筛选条件和第几个。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` `int` | `''` | 用于筛选节点的  查询语法，为`int`类型时`index`参数无效 |
| `index` | `int` | `1` | 查询结果中的第几个，从`1`开始，可输入负数表示倒数 |
| `timeout` | `float` | `None` | 查找超时时间（秒），为`None`时使用页面超时设置，s 模式下无效 |
| `ele_only` | `bool` | `True` | 是否只查找元素，为`False`时把文本、注释节点也纳入查找范围 |

| 返回类型 | 说明 |
| --- | --- |
| `str` | 获取非元素节点时返回字符串 |
| `ChromiumElement` | d 模式下的元素对象 |
| `SessionElement` | s 模式下的元素对象 |
| `NoneElement` | 未获取到结果时 |

**示例：**

```
# 获取 ele1 前面第一个兄弟元素  
ele2 = ele1.prev()  
  
# 获取 ele1 前面第 3 个兄弟元素  
ele2 = ele1.prev(3)  
  
# 获取 ele1 前面第 3 个 div 兄弟元素  
ele2 = ele1.prev(3, 'tag:div')  
  
# 获取 ele1 前面第一个文本节点的文本  
txt = ele1.prev(1, 'xpath:text()')
```

---

🔸 `prevs()`

此方法返回当前元素前面全部符合条件的同级节点组成的列表，可用查询语法筛选。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` | `''` | 用于筛选节点的查询语法 |
| `timeout` | `float` | `None` | 查找超时时间（秒），为`None`时使用页面超时设置，s 模式下无效 |
| `ele_only` | `bool` | `True` | 是否只查找元素，为`False`时把文本、注释节点也纳入查找范围 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElementsList` | d 模式结果列表 |
| `SessionElementsList` | s 模式结果列表 |

**示例：**

```
# 获取 ele1 前面所有兄弟元素  
eles = ele1.prevs()  
  
# 获取 ele1 前面所有 div 兄弟元素  
divs = ele1.prevs('tag:div')
```

---

### 📌 在后面文档中查找节点[​](#-在后面文档中查找节点 "📌 在后  面文档中查找节点的直接链接")

🔸 `after()`

此方法返回当前元素后面的某一个节点，可指定筛选条件和第几个。查找范围不限同级节点，而是整个 DOM 文档。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` `int` | `''` | 用于筛选节点的查询语法，为`int`类型时`index`参数无效 |
| `index` | `int` | `1` | 查询结果中的第几个，从`1`开始，可输入负数表示倒数 |
| `timeout` | `float` | `None` | 查找超时时间（秒），为`None`时使用页面超时设置，s 模式下无效 |
| `ele_only` | `bool` | `True` | 是否只查找元素，为`False`时把文本、注释节点也纳入查找范围 |

| 返回类型 | 说明 |
| --- | --- |
| `str` | 获取非元素节点时返回字符串 |
| `ChromiumElement` | d 模式下的元素对象 |
| `SessionElement` | s 模式下的元素对象 |
| `NoneElement` | 未获取到结果时 |

**示例：**

```
# 获取 ele1 后面第 3 个元素  
ele2 = ele1.after(index=3)  
  
# 获取 ele1 后面第 3 个 div 元素  
ele2 = ele1.after('tag:div', 3)  
  
# 获取 ele1 后面第一个文本节点的文本  
txt = ele1.after('xpath:text()', 1)
```

---

🔸 `afters()`

此方法返回当前元素后面符合条件的全部节点组成的列表，可用查询语法筛选。查找范围不限同级节点，而是整个 DOM 文档。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` | `''` | 用于筛选节点的查询语法 |
| `timeout` | `float` | `None` | 查找超时时间（秒），为`None`时使用页面超时设置，s 模式下无效 |
| `ele_only` | `bool` | `True` | 是否只查找元素，为`False`时把文本、注释节点也纳入查找范围 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElementsList` | d 模式结果列表 |
| `SessionElementsList` | s 模式结果列表 |

**示例：**

```
# 获取 ele1 后所有元素  
eles = ele1.afters()  
  
# 获取 ele1 前面所有 div 元素  
divs = ele1.afters('tag:div')
```

---

### 📌 在前面文档中查找节点[​](#-在前面文档中查找节点 "📌 在前面文档中查找节点的直  接链接")

🔸 `before()`

此方法返回当前元素前面的某一个符合条件的节点，可指定筛选条件和第几个。查找范围不限同级节点，而是整个 DOM 文档。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` `int` | `''` | 用于筛选节点的查询语法，为`int`类型时`index`参数无效 |
| `index` | `int` | `1` | 查询结果中的第几个，从`1`开始，可输入负数表示倒数 |
| `timeout` | `float` | `None` | 查找超时时间（秒），为`None`时使用页面超时设置，s 模式下无效 |
| `ele_only` | `bool` | `True` | 是否只查找元素，为`False`时把文本、注释节点也纳入查找范围 |

| 返回类型 | 说明 |
| --- | --- |
| `str` | 获取非元素节点时返回字符串 |
| `ChromiumElement` | d 模式下的元素对象 |
| `SessionElement` | s 模式下的元素对象 |
| `NoneElement` | 未获取到结果时 |

**示例：**

```
# 获取 ele1 前面第 3 个元素  
ele2 = ele1.before(3)  
  
# 获取 ele1 前面第 3 个 div 元素  
ele2 = ele1.before('tag:div', 3)  
  
# 获取 ele1 前面第一个文本节点的文本  
txt = ele1.before('xpath:text()', 1)
```

---

🔸 `befores()`

此方法返回当前元素前面全部符合条件的节点组成的列表，可用查询语法筛选。查找范围不限同级节点，而是整个 DOM 文档。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` | `''` | 用于筛选节点的查询语法 |
| `timeout` | `float` | `None` | 查找超时时间（秒），为`None`时使用页面超时设置，s 模式下无效 |
| `ele_only` | `bool` | `True` | 是否只查找元素，为`False`时把文本、注释节点也纳入查找范围 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElementsList` | d 模式结果列表 |
| `SessionElementsList` | s 模式结果列表 |

**示例：**

```
# 获取 ele1 前面所有元素  
eles = ele1.befores()  
  
# 获取 ele1 前面所有 div 元素  
divs = ele1.befores('tag:div')
```

---

  

✅️️ 基于视觉相对定位[​](#️️-基于视觉相对定位 "✅️️ 基于视觉相对定位的直接链接")
-------------------------------------------------

以下方法可以以某元素为基准，向不同方向或指定偏移量获取元素。

只有浏览器模式支持这类定位方式。

只能获取可见的元素（不论是否在视口内），不能获取被遮挡的。

### 📌 `east()`[​](#-east "-east的直接链接")

此方法用于获取一个在当前元素右边的元素。

`loc_or_pixel`参数可用定位符指定筛选条件，定位符只支持`str`格式，且不支持 xpath 和 css 方式。

用`index`参数可指定获取第几个结果。如果`loc_or_pixel`为`None`，获取第若干个元素。

`loc_or_pixel`为`int`格式时，直接获取元素右边这个距离的元素，此时`index`参数无效。距离从右边框开始计算。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `loc_or_pixel` | `str` `int` | `None` | 定位符或距离（像素） |
| `index` | `int` | `1` | 第几个，从1开始，`loc_or_pixel`为`int`格式时无效 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 找到的元素对象 |
| `NoneElement` | 未获取到结果时 |

---

### 📌 `west()`[​](#-west "-west的直接链接")

此方法用于获取一个在当前元素左边的元素。

`loc_or_pixel`参数可用定位符指定筛选条件，定位符只支持`str`格式，且不支持 xpath 和 css 方式。

用`index`参数可指定获取第几个结果。如果`loc_or_pixel`为`None`，获取第若干个元素。

`loc_or_pixel`为`int`格式时，直接获取元素左边这个距离的元素，此时`index`参数无效。距离从左边框开始计算。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `loc_or_pixel` | `str` `int` | `None` | 定位符或距离（像素） |
| `index` | `int` | `1` | 第几个，从1开始，`loc_or_pixel`为`int`格式时无效 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 找到的元素对象 |
| `NoneElement` | 未获取到结果时 |

---

### 📌 `south()`[​](#-south "-south的直接链接")

此方法用于获取一个在当前元素下边的元素。

`loc_or_pixel`参数可用定位符指定 筛选条件，定位符只支持`str`格式，且不支持 xpath 和 css 方式。

用`index`参数可指定获取第几个结果。如果`loc_or_pixel`为`None`，获取第若干个元素。

`loc_or_pixel`为`int`格式时，直接获取元素下边这个距离的元素，此时`index`参数无效。距离从下边框开始计算。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `loc_or_pixel` | `str` `int` | `None` | 定位符或距离（像素） |
| `index` | `int` | `1` | 第几个，从1开始，`loc_or_pixel`为`int`格式时无效 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 找到的元素对象 |
| `NoneElement` | 未获取到结果时 |

---

### 📌 `north()`[​](#-north "-north的直接链接")

此方法用于获取一个在当前元素上边的元素。

`loc_or_pixel`参数可用定位符指定筛选条件，定位符只支持`str`格式，且不支持 xpath 和 css 方式。

用`index`参数可指定获取第几个结果。如果`loc_or_pixel`为`None`，获取第若干个元素。

`loc_or_pixel`为`int`格式时，直接获取元素上边这个距离的元素，此时`index`参数无效。距离从上边框开始计算。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `loc_or_pixel` | `str` `int` | `None` | 定位符或距离（像素） |
| `index` | `int` | `1` | 第几个，从1开始，`loc_or_pixel`为`int`格式时无效 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 找到的元素对象 |
| `NoneElement` | 未获取到结果时 |

---

### 📌 `offset()`[​](#-offset "-offset的直接链接")

此方法用于获取相对于元素左上角指定偏移量的一个元素。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `offset_x` | `int` | 必填 | 横坐标偏移量（像素），向右为正 |
| `offset_y` | `int` | 必填 | 纵坐标偏移量（像素），向下为正 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 找到的元素对象 |
| `NoneElement` | 未获取到结果时 |

---

### 📌 `over()`[​](#-over "-over的直接链接")

此方法用于获取覆盖在本元素上最上层的元素。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` | `None` | 等待元素出现的超时时间（秒），为`None`使用页面`timeout`设置值 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 找到的元素对象 |
| `NoneElement` | 未获取到结果时 |

[上一页

🔦 页面或元素内查找](/browser_control/get_elements/find_in_object)[下一页

🔦 行为模式](/browser_control/get_elements/behavior)

* [✅️️ 基于 DOM 相对定位](#️️-基于-dom-相对定位)
  + [📌 获取父级元素](#-获取父级元素)
  + [📌 获取直接子节点](#-获取直接子节点)
  + [📌 获取后面的同级节点](#-获取后面的同级节点)
  + [📌 获取前面的同级节点](#-获取前面的同级节点)
  + [📌 在后面文档中查找节点](#-在后面文档中查找节点)
  + [📌 在前面文档中查找节点](#-在前面文档中查找节点)
* [✅️️ 基于视觉相对定位](#️️-基于视觉相对定位)
  + [📌 `east()`](#-east)
  + [📌 `west()`](#-west)
  + [📌 `south()`](#-south)
  + [📌 `north()`](#-north)
  + [📌 `offset()`](#-offset)
  + [📌 `over()`](#-over)

* 🚀 控制浏览器
* 🔎 查找元素
* 🔦 语法速查表

本页总览

🔦 语法速查表
=======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

✅️ 定位语法[​](#️-定位语法 "✅️ 定位语法的直接链接")
----------------------------------

### 📌 基本用法[​](#-基本用法 "📌 基本用法的直接链接")

| 写法 | 精确匹配 | 模糊匹配 | 匹配开头 | 匹配结尾 | 说明 |
| --- | --- | --- | --- | --- | --- |
| `@属性名` | `@属性名=` | `@属性名:` | `@属性名^` | `@属性名$` | 按某个属性查找 |
| `@!属性名` | `@!属性名=` | `@!属性名:` | `@!属性名^` | `@!属性名$` | 查找属性不符合指定条件的元素 |
| `text` | `text=` | `text:`或不写 | `text^` | `text$` | 按某个文本查找 |
| `@text()` | `@text()=` | `@text():` | `text()^` | `text()$` | `text`与`@`或`@@`配合使用时改为`text()`，常用于多条件匹配 |
| `tag` | `tag=`或`tag:` | 无 | 无 | 无 | 查找某个类型的元素 |
| `@tag()` | `@tag()=`或`@tag():` | 无 | 无 | 无 | 组合使用时查找某个类型的元素 |
| `xpath` | `xpath=`或`xpath:` | 无 | 无 | 无 | 用 xpath 方式查找元素 |
| `css` | `css=`或`css:` | 无 | 无 | 无 | 用 css selector 方式查找元素 |

---

### 📌 组合用法[​](#-组合用法 "📌 组合用法的直接链接")

| 写法 | 说明 |
| --- | --- |
| `@@属性1@@属性2` | 匹配属性同时符 合多个条件的元素 |
| `@@属性1@!属性2` | 多属性匹配与否定匹配同时使用 |
| `@|属性1@|属性2` | 匹配属性至符合多个条件中一的元素 |
| `tag:元素名@属性名` | `tag`与属性匹配共同使用 |
| `tag:元素名@@属性1@@属性2` | `tag`与多属性匹配共同使用 |
| `tag:元素名@|属性1@|属性2` | `tag`与多属性匹配共同使用 |
| `tag:元素名@@text()=文本@@属性` | `tag`与文本和属性匹配共同使用 |

---

### 📌 简化写法[​](#-简化写法 "📌 简化写法的直接链接")

| 原写法 | 简化写法 | 精确匹配 | 模糊匹配 | 匹配开头 | 匹配结尾 | 备注 |
| --- | --- | --- | --- | --- | --- | --- |
| `@id` | `#` | `#`或`#=` | `#:` | `#^` | `#$` | 简化写法只能单独使用 |
| `@class` | `.` | `.`或`.=` | `.:` | `.^` | `.$` | 简化写法只能单独使用 |
| `tag` | `t` | `t:`或`t=` | 无 | 无 | 无 | 只能用在句首 |
| `@tag()` | `@t()` | `@t():`或`@t()=` | 无 | 无 | 无 | 可作为属性组合使用 |
| `text` | `tx` | `tx=` | `tx:`或不写 | `tx^` | `tx$` | 无标签时使用模糊匹配文本 |
| `@text()` | `@tx()` | `@tx()=` | `@tx():` | `@tx()^` | `@tx()$` | 可作为属性组合使用 |
| `xpath` | `x` | `x:`或`x=` | 无 | 无 | 无 | 只能单独使用 |
| `css` | `c` | `c:`或`c=` | 无 | 无 | 无 | 只能单独使用 |

---

✅️ 相对定位[​](#️-相对定位 "✅️ 相对定位的直接链接")
----------------------------------

| 方法 | 说明 |
| --- | --- |
| `parent()` | 查找当前元素某一级父元素 |
| `child()` | 查找当前元素的一个直接子节点 |
| `children()` | 查找当前元素全部符合条件的直接子节点 |
| `next()` | 查找当前元素之后第一个符合条件的兄弟节点 |
| `nexts()` | 查找当前元素之后所有符合条件的兄弟节点 |
| `prev()` | 查找当前元素之前第一个符合条件的兄弟节点 |
| `prevs()` | 查找当前元素之前所有符合条件的兄弟节点 |
| `after()` | 查找文档中当前元素之后第一个符合条件的节点 |
| `afters()` | 查找文档中当前元素之后所有符合条件的节点 |
| `before()` | 查找文档中当前元素之前第一个符合条件的节点 |
| `befores()` | 查找文档中当前元素之前所有符合条件的节点 |

  

✅️ iframe 和 shadow root[​](#️-iframe-和-shadow-root "✅️ iframe 和 shadow root的直接链接")
----------------------------------------------------------------------------------

| 方法 | 简化写法 | 说明 | 备注 |
| --- | --- | --- | --- |
| `get_frame()` | 无 | 在页面中查找一个`<iframe>`元素 | 只有页面对象有此方法 |
| `shadow_root` | `sr` | 获取当前元素内的 shadow root 对象 | 只有元素对象有此属性 |

---

✅️ 特殊字符对照表[​](#️-特殊字符对照表 "✅️ 特殊字符对照表的直接链接")
-------------------------------------------

要匹配的文本中如包含特殊字符（如`'&nbsp;'`、`'&gt;'`），需将特殊字符转为十六进制，对照表如下：

| 特殊符号 | 命名实体 | 编码 | 特殊符号 | 命名实体 | 编码 | 特殊符号 | 命名实体 | 编码 |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| `Α` | `&Alpha;` | \u0391 | `Β` | `&Beta;` | \u0392 | `Γ` | `&Gamma;` | \u0393 |
| `Δ` | `&Delta;` | \u0394 | `Ε` | `&Epsilon;` | \u0395 | `Ζ` | `&Zeta;` | \u0396 |
| `Η` | `&Eta;` | \u0397 | `Θ` | `&Theta;` | \u0398 | `Ι` | `&Iota;` | \u0399 |
| `Κ` | `&Kappa;` | \u039A | `Λ` | `&Lambda;` | \u039B | `Μ` | `&Mu;` | \u039C |
| `Ν` | `&Nu;` | \u039D | `Ξ` | `&Xi;` | \u039E | `Ο` | `&Omicron;` | \u039F |
| `Π` | `&Pi;` | \u03A0 | `Ρ` | `&Rho;` | \u03A1 | `Σ` | `&Sigma;` | \u03A3 |
| `Τ` | `&Tau;` | \u03A4 | `Υ` | `&Upsilon;` | \u03A5 | `Φ` | `&Phi;` | \u03A6 |
| `Χ` | `&Chi;` | \u03A7 | `Ψ` | `&Psi;` | \u03A8 | `Ω` | `&Omega;` | \u03A9 |
| `α` | `&alpha;` | \u03B1 | `β` | `&beta;` | \u03B2 | `γ` | `&gamma;` | \u03B3 |
| `δ` | `&delta;` | \u03B4 | `ε` | `&epsilon;` | \u03B5 | `ζ` | `&zeta;` | \u03B6 |
| `η` | `&eta;` | \u03B7 | `θ` | `&theta;` | \u03B8 | `ι` | `&iota;` | \u03B9 |
| `κ` | `&kappa;` | \u03BA | `λ` | `&lambda;` | \u03BB | `μ` | `&mu;` | \u03BC |
| `ν` | `&nu;` | \u03BD | `ξ` | `&xi;` | \u03BE | `ο` | `&omicron;` | \u03BF |
| `π` | `&pi;` | \u03C0 | `ρ` | `&rho;` | \u03C1 | `ς` | `&sigmaf;` | \u03C2 |
| `σ` | `&sigma;` | \u03C3 | `τ` | `&tau;` | \u03C4 | `υ` | `&upsilon;` | \u03C5 |
| `φ` | `&phi;` | \u03C6 | `χ` | `&chi;` | \u03C7 | `ψ` | `&psi;` | \u03C8 |
| `ω` | `&omega;` | \u03C9 | `ϑ` | `&thetasym;` | \u03D1 | `ϒ` | `&upsih;` | \u03D2 |
| `ϖ` | `&piv;` | \u03D6 | `•` | `&bull;` | \u2022 | `…` | `&hellip;` | \u2026 |
| `′` | `&prime;` | \u2032 | `″` | `&Prime;` | \u2033 | `‾` | `&oline;` | \u203E |
| `⁄` | `&frasl;` | \u2044 | `℘` | `&weierp;` | \u2118 | `ℑ` | `&image;` | \u2111 |
| `ℜ` | `&real;` | \u211C | `™` | `&trade;` | \u2122 | `ℵ` | `&alefsym;` | \u2135 |
| `←` | `&larr;` | \u2190 | `↑` | `&uarr;` | \u2191 | `→` | `&rarr;` | \u2192 |
| `↓` | `&darr;` | \u2193 | `↔` | `&harr;` | \u2194 | `↵` | `&crarr;` | \u21B5 |
| `⇐` | `&lArr;` | \u21D0 | `⇑` | `&uArr;` | \u21D1 | `⇒` | `&rArr;` | \u21D2 |
| `⇓` | `&dArr;` | \u21D3 | `⇔` | `&hArr;` | \u21D4 | `∀` | `&forall;` | \u2200 |
| `∂` | `&part;` | \u2202 | `∃` | `&exist;` | \u2203 | `∅` | `&empty;` | \u2205 |
| `∇` | `&nabla;` | \u2207 | `∈` | `&isin;` | \u2208 | `∉` | `&notin;` | \u2209 |
| `∋` | `&ni;` | \u220B | `∏` | `&prod;` | \u220F | `∑` | `&sum;` | \u2212 |
| `−` | `&minus;` | \u2212 | `∗` | `&lowast;` | \u2217 | `√` | `&radic;` | \u221A |
| `∝` | `&prop;` | \u221D | `∞` | `&infin;` | \u221E | `∠` | `&ang;` | \u2220 |
| `∧` | `&and;` | \u22A5 | `∨` | `&or;` | \u22A6 | `∩` | `&cap;` | \u2229 |
| `∪` | `&cup;` | \u222A | `∫` | `&int;` | \u222B | `∴` | `&there4;` | \u2234 |
| `∼` | `&sim;` | \u223C | `≅` | `&cong;` | \u2245 | `≈` | `&asymp;` | \u2245 |
| `≠` | `&ne;` | \u2260 | `≡` | `&equiv;` | \u2261 | `≤` | `&le;` | \u2264 |
| `≥` | `&ge;` | \u2265 | `⊂` | `&sub;` | \u2282 | `⊃` | `&sup;` | \u2283 |
| `⊄` | `&nsub;` | \u2284 | `⊆` | `&sube;` | \u2286 | `⊇` | `&supe;` | \u2287 |
| `⊕` | `&oplus;` | \u2295 | `⊗` | `&otimes;` | \u2297 | `⊥` | `&perp;` | \u22A5 |
| `⋅` | `&sdot;` | \u22C5 | `⌈` | `&lceil;` | \u2308 | `⌉` | `&rceil;` | \u2309 |
| `⌊` | `&lfloor;` | \u230A | `⌋` | `&rfloor;` | \u230B | `◊` | `&loz;` | \u25CA |
| `♠` | `&spades;` | \u2660 | `♣` | `&clubs;` | \u2663 | `♥` | `&hearts;` | \u2665 |
| `♦` | `&diams;` | \u2666 |  | `&nbsp;` | \u00A0 | `¡` | `&iexcl;` | \u00A1 |
| `¢` | `&cent;` | \u00A2 | `£` | `&pound;` | \u00A3 | `¤` | `&curren;` | \u00A4 |
| `¥` | `&yen;` | \u00A5 | `¦` | `&brvbar;` | \u00A6 | `§` | `&sect;` | \u00A7 |
| `¨` | `&uml;` | \u00A8 | `©` | `&copy;` | \u00A9 | `ª` | `&ordf;` | \u00AA |
| `«` | `&laquo;` | \u00AB | `¬` | `&not;` | \u00AC | `­` | `&shy;` | \u00AD |
| `®` | `&reg;` | \u00AE | `¯` | `&macr;` | \u00AF | `°` | `&deg;` | \u00B0 |
| `±` | `&plusmn;` | \u00B1 | `²` | `&sup2;` | \u00B2 | `³` | `&sup3;` | \u00B3 |
| `´` | `&acute;` | \u00B4 | `µ` | `&micro;` | \u0012 | `"` | `&quot;` | \u0022 |
| `<` | `&lt;` | \u003C | `>` | `&gt;` | \u003E | `'` |  | \u0027 |

[上一页

🔦 简化写法](/browser_control/get_elements/simplify)[下一页

🛰️ 元素交互](/browser_control/ele_operation)

* [✅️ 定位语法](#️-定位语法)
  + [📌 基本用法](#-基本用法)
  + [📌 组合用法](#-组合用法)
  + [📌 简化写法](#-简化写法)
* [✅️ 相对定位](#️-相对定位)
* [✅️ iframe 和 shadow root](#️-iframe-和-shadow-root)
* [✅️ 特殊字符对照表](#️-特殊字符对照表)

* 🚀 控制浏览器
* 🔎 查找元素
* 🔦 简化写法

本页总览

🔦 简化写法
======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

为进一步精简代码，定位语法都可以用简化形式来表示，使语句更短，链式操作时更清晰。

✅️ 定位符语法简化[​](#️-定位符语法简化 "✅️ 定位符语法简化的直接链接")
-------------------------------------------

* 定位语法都有其简化形式
* 页面和元素对象都实现了`__call__()`方法，所以`page.ele('****')`可简化为`page('****')`
* 查找方法都支持链式操作

示例：

```
# 查找tag为div的元素  
ele = tab.ele('tag:div')  # 原写法  
ele = tab('t:div')  # 简化写法  
  
# 用xpath查找元素  
ele = tab.ele('xpath://****')  # 原写法  
ele = tab('x://****')  # 简化写法  
  
# 查找text为'something'的元素  
ele = tab.ele('text=something')  # 原写法  
ele = tab('tx=something')  # 简化写法
```

简化写法对应列表

| 原写法 | 简化写法 | 说明 |
| --- | --- | --- |
| `@id` | `#` | 表示 id 属性，简化写法只在语句最前面且单独使用时生效 |
| `@class` | `.` | 表示 class 属性，简化写法只在语句最前面且单独使用时生效 |
| `text` | `tx` | 按文本匹配 |
| `@text()` | `@tx()` | 按文本查找，与 @ 或 @@ 配合使用时 |
| `tag` | `t` | 按标签类型匹配 |
| `@tag()` | `@t()` | 按元素名查找，与 @ 或 @@ 配合使用时 |
| `xpath` | `x` | 用 xpath 方式查找元素 |
| `css` | `c` | 用 css selector 方式查找元素 |

---

✅️ shadow root 简化[​](#️-shadow-root-简化 "✅️ shadow root 简化的直接链接")
----------------------------------------------------------------

一般获取元素的 shadow root 元素，用`ele.shadow_root`属性。

由于此属性经常用于大量链式操作，名字太长影响可读性，因此可简化为`ele.sr`

**示例：**

```
txt = ele.sr('t:div').text
```

---

  

✅️ 相对定位参数简化[​](#️-相对定位参数简化 "✅️ 相对定位参数简化的直接链接")
----------------------------------------------

相对定位时，有时需要获取当前元素后某个元素，而不关心该元素是什么类型，一般是这样写：`ele.next(index=2)`。

但有一种简化的写法，可以直接写作`ele.next(2)`。

当第一个参数`filter_loc`接收数字时，会自动将其视作序号，替代`index`参数。因此书写可以稍微精简一些。

**示例：**

```
ele2 = ele1.parent(2)  
ele2 = ele1.next(2)('tx=****')  
ele2 = ele1.before(2)  
# 如此类推
```

[上一页

🔦 在结果列表中筛选](/browser_control/get_elements/filter)[下一页

🔦 语法速查表](/browser_control/get_elements/sheet)

* [✅️ 定位符语法简化](#️-定位符语法简化)
* [✅️ shadow root 简化](#️-shadow-root-简化)
* [✅️ 相对定位参数简化](#️-相对定位参数简化)

* 🚀 控制浏览器
* 🔎 查找元素
* 🔦 定位语法

本页总览

🔦 定位语法
======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

定位语法用于指明以哪种方式去查找指定元素，语法简洁明了，熟练使用可大幅提高程序可读性。

所有涉及获取元素的操作都可以使用定位语法，如`ele()`、`actions.move_to()`、`wait.eles_loaded()`、`get_frame()`等等。

定位语法用于简化代码，提高可读性，但并不覆盖所有复杂场景。很复杂的场景可直接用 xpath 查找。

以下使用这个页面进行讲解。

```
<html>  
<body>  
<div id="one">  
    <p class="p_cls" id="row1" data="a">第一行</p>  
    <p class="p_cls" id="row2" data="b">第二行</p>  
    <p class="p_cls">第三行</p>  
</div>  
<div id="two">  
    第二个div  
</div>  
</body>  
</html>
```

✅️️ 基本概念[​](#️️-基本概念 "✅️️ 基本概念的直接链接")
-------------------------------------

几乎所有查找方法都是基于元素属性进行。

元素属性包括以下三种：

| 写法 | 说明 | 示例 |
| --- | --- | --- |
| `@tag()` | 标签名 | 即`<div id="one">`中的`div` |
| `@****` | 标签体中的属性 | 如`<div id="one">`中的`id`，写作`'@id'` |
| `@text()` | 元素文本 | 即`<p class="p_cls">第三行</p>`中的`第三行` |

查找语法就是按需要对这三种属性进行组合，达到查找指定元素的目的。

说明

`@tag()`和`@text()`后面加上`'()'`，是为了避免与普通元素冲突（如`<div text="abc">`）。

### 📌 简单示例[​](#-简单示例 "📌 简单示例的直接链接")

```
tab.ele('@id=one')  # 获取第一个id为one的元素  
tab.ele('@tag()=div')  # 获取第一个div元素  
tab.ele('@text()=第一行')  # 获取第一个文本为“第一行”的元素
```

---

✅️️ 基本逻辑[​](#️️-基本逻辑 "✅️️ 基本逻辑的直接链接")
-------------------------------------

### 📌 单属性匹配符 `@`[​](#-单属性匹配符- "-单属性匹配符-的直接链接")

单个`@`在只以一个属性作为匹配条件时使用，以`'@'`开头，后面跟属性名称。

上面简单示例中就是这种方式：`tab.ele('@id=one')`。

如果`@`后面只有属性名而没有属性值，查找有这个属性的元素，如`tab.ele('@id')`。

注意

如果属性中包含特殊字符（如包含`@`），用这个方式不能正确匹配到，需使用 css selector 方式查找。且特殊字符要用`\`转义。

---

### 📌 多属性与匹配符 `@@`[​](#-多属性与匹配符- "-多属性与匹配符-的直接链接")

当需要多个条件同时确定一个元素时，每个属性用`'@@'`开头。

注意

* 匹配文本或属性中出现`@@`、`@|`、`@!`时，不能使用多属性匹配，需改用 xpath 的方式。
* 如果属性中包含特殊字符（如包含`@`），用这个方式不能正确匹配到，需使用 css selector 方式查找。且特殊字符要用`\`转义。

**示例：**

```
ele = tab.ele('@@class=p_cls@@text()=第三行')  # 查找class为p_cls且文本为“第三行”的元素
```

---

### 📌 多属性或匹配符 `@|`[​](#-多属性或匹配符- "-多属性或匹配符-的直接链接")

当需要以或关系条件查找元素时，每个属性用`'@|'`开头。

注意

* 匹配文本或属性中出现`@@`、`@|`、`@!`时，不能使用多属性匹配，需改用 xpath 的方式。
* 如果属性中包含特殊字符（如包含`@`），用这个方式不能正确匹配到，需使用 css selector 方式查找。且特殊字符要用`\`转义。

**示例：**

```
eles = tab.eles('@|id=row1@|id=row2')  # 查找所有id为row1或id为row2的元素
```

---

### 📌 否定匹配符 `@!`[​](#-否定匹配符- "-否定匹配符-的直接链接")

用于否定某个条件。

如果`@!`后面只有属性名而没有属性值，查找没有这个属性的元素。

**示例：**

```
ele = tab.ele('@!id=one')  # 获取第一个id  不等于“one”的元素  
ele = tab.ele('@!class')  # 匹配没有class属性的元素
```

---

### 📌 混合使用[​](#-混合使用 "📌 混合使用的直接链接")

`@@`和`@|`不能同时出现的查找语句中，即一个查找语句只能是与关系或者或关系。

`@!`则可与两者混合使用。混用时，与还是或关系视`@@`还是`@|`而定。

说明

当语句中有多个`tag()`时，如果全部都没有被`@!`修饰，它们是与关系；如有任一个被`@!`修饰，它们是或关系。
`tag()`与其他属性之间是与关系。

**示例：**

```
# 匹配class等于p_cls且id不等于row1的元素  
tab.ele('@@class=p_cls@!id=row1')  
  
# 匹配class等于p_cls或id不等于row1的元素  
tab.ele('@|class=p_cls@!id=row1')
```

---

  

✅️️ 匹配模式[​](#️️-匹配模式 "✅️️ 匹配模式的直接链接")
-------------------------------------

匹配模式指某个查询中匹配条件的方式，有精确匹配、模糊匹配、匹配开头、匹配结尾四种。

说明

`tag()`属性无论用哪种匹配模式，都会视作`=`。

### 📌 精确匹配 `=`[​](#-精确匹配- "-精确匹配-的直接链接")

表示精确匹配，匹配完全符合的文本或属性。

```
ele = tab.ele('@id=row1')  # 获取id属性为'row1'的元素
```

---

### 📌 模糊匹配 `:`[​](#-模糊匹配- "-模糊匹配-的直接链接")

表示模糊匹配，匹配含有指定字符串的文本或属性。

```
ele = tab.ele('@id:ow')  # 获取id属性包含'ow'的元素
```

---

### 📌 匹配开头 `^`[​](#-匹  配开头- "-匹配开头-的直接链接")

表示匹配开头，匹配开头为指定字符串的文本或属性。

```
ele = tab.ele('@id^row')  # 获取id属性以'row'开头的元素
```

---

### 📌 匹配结尾 `$`[​](#-匹配结尾- "-匹配结尾-的直接链接")

 表示匹配结尾，匹配结尾为指定字符串的文本或属性。

```
ele = tab.ele('@id$w1')  # 获取id属性以'w1'结尾的元素
```

---

✅️️ 常用语法[​](#️️-常用语法 "✅️️ 常用语法的直接链接")
-------------------------------------

基于上述基本逻辑，本库提供了一些更易于使用和阅读的语法。

### 📌 id 匹配符 `#`[​](#-id-匹配符- "-id-匹配符-的直接链接")

用于匹配`id`属性，**只在语句最前面且单独使用时生效**。相当于单属性查找`@id=****`。

可与匹配模式配合使用。

```
ele = tab.ele('#one')  # 查找id为one的元素  
ele = tab.ele('#=one')  # 和上面一行一致  
ele = tab.ele('#:ne')  # 查找id属性包含ne的元素  
ele = tab.ele('#^on')  # 查找id属性以on开头的元素  
ele = tab.ele('#$ne')  # 查找id属性以ne结尾的元素
```

---

### 📌 class 匹配符 `.`[​](#-class-匹配符- "-class-匹配符-的直接链接")

用于匹配`class`属性，**只在语句最前面且单独使用时生效**，相当于单属性查找`@class=****`。

可配合匹配模式使用。

说明

在面对多个 class 的元素时，DrissionPage 与 selenium 处理方式不一样，无需将空格替换成`'.'`。
而是将整个 class 视作普通字符串，空格视作普通字符对待，会比较直观。

```
ele = tab.ele('.p_cls')  # 查找class属性为p_cls的元素  
ele = tab.ele('.=p_cls')  # 与上一行一致  
ele = tab.ele('.:_cls')  # 查找class属性包含_cls的元素  
ele = tab.ele('.^p_')  # 查找class属性以p_开头的元素  
ele = tab.ele('.$_cls')  # 查找class属性以_cls结尾的元素
```

---

###    📌 文本匹配符 `text`[​](#-文本匹配符-text "-文本匹配符-text的直接链接")

用于匹配元素文本。**只在语句最前面且单独使用时生效**，相当于单属性查找`@text()=****`。

可配合匹配模式使用。

如果元素内有多个直接的文本节点，精确查找时可匹配所有文本节点拼成的字符串，模糊查找时可匹配每个文本节点。

如果查找语句没有任何本节介绍的匹配符，默认模糊匹配文本。即`ele('第三行')`相当于`ele('text:第三行')`。

注意

如果要匹配的文本包含特殊字符（如`'&nbsp;'`、`'&gt;'`），需将其转换为十六进制形式，详见《语法速查表》一节。

```
ele = tab.ele('text=第二行')  # 查找文本为“第二行”的元素  
ele = tab.ele('text:第二')  # 查找文本包含“第二”的元素  
ele = tab.ele('第二')  # 与上一行一致  
ele = tab.ele('第\u00A0二')  # 匹配包含&nbsp;文本的元素，需将&nbsp;转为\u00A0
```

Tips

若要查找的文本包含`text:` ，可下面这样写，即第一个`text:` 为关键字，第二个是要查找的内容：

```
ele2 = tab.ele('text:text:')
```

---

### 📌 类型匹配符 `tag`[​](#-类型匹配符-tag "-类型匹配符-tag的直接链接")

用于匹配某类型元素。**只在语句最前面且单独使用时生效**，相当于单属性查找`@tag()=****`。

可与单属性查找或多属性配合使用。`tag:`与`tag=`效果一致，没有`tag^`和`tag$`语法。

```
ele = tab.ele('tag:div')  # 查找第一个div元素  
ele = tab.ele('tag:p@class=p_cls')  # 与单属性查找配合使用  
ele = tab.ele('tag:p@@class=p_cls@@text()=第二行')  # 与多属性查找配合使用
```

注意

`tag:div@text():abc` 和 `tag:div@@text():abc` 是有区别的，前者只在`div`的直接文本节点搜索，后者搜索`div`的整个内部。

---

### 📌 css selector 匹配符 `css`[​](#-css-selector-匹配符-css "-css-selector-匹配符-css的直接链接")

表示用 css selector 方式查找元素。**只在语句最前面且单独使用时生效**。

`css:`与`css=`效果一致  ，没有`css^`和`css$`语法。

```
ele = tab.ele('css:.div')  # 查找 div 元素  
ele = tab.ele('css:>div')  # 查找 div 子元素元素，这个写法是本库特有，原生不支持
```

---

### 📌 xpath 匹配符 `xpath`[​](#-xpath-匹配符-xpath "-xpath-匹配符-xpath的直接链接")

表示用 xpath 方式查找元素。**只在语句最前面且单独使用时生效**。

`xpath:`与`xpath=`效果一致，没有`xpath^`和`xpath$`语法。

Tips

**元素对象**的`ele()`支持完整的 xpath 语法，如能使用 xpath 直接获取元素属性（字符串类型）。

```
ele2 = ele1.ele('xpath:.//div')  # 查找后代中第一个 div 元素  
ele2 = ele1.ele('xpath://div')  # 和上面一行一样，查找元素的后代时，// 前面的 . 可以省略  
ele_class_str = ele1.ele('xpath://div/@class')  # 使用xpath获取div元素的class属性（页面元素无此功能）
```

说明

查找元素的后代时，selenium 原生代码要求 xpath 前面必须加`.`，否则会变成在全个页面中查找。
作者觉得这个设计是画蛇添足，既然已经通过元素查找了，自然应该只查找这个元素内部的元素。
所以，用 xpath 在元素下查找时，最前面`//`或`/`前面的`.`可以省略。

---

### 📌 selenium 的 loc 元组[​](#-selenium-的-loc-元组 "📌 selenium 的 loc 元组的直接链接")

查找方法能直接接收 selenium 原生定位元组进行查找，便于项目迁移。

```
from DrissionPage.common import By  
  
# 查找id为one的元素  
loc1 = (By.ID, 'one')  
ele = tab.ele(loc1)  
  
# 按 xpath 查找  
loc2 = (By.XPATH, '//p[@class="p_cls"]')  
ele = tab.ele(loc2)
```

---

✅️️ `@@text()`的技巧[​](#️️-text的技巧 "️️-text的技巧的直接链接")
---------------------------------------------------

值得一提的是，`text()`配合`@@`或`@|`能实现一种很便利的按查找方式。

网页种经常会出现元素和文本混排的情况，比如：

```
<li class="explore-categories__item">  
    <a href="/explore/new-tech" class="">  
        <i class="explore"></i>  
        前沿技术  
    </a>  
</li>  
<li class="explore-categories__item">  
    <a href="/explore/program-develop" class="">  
        <i class="explore"></i>  
        程序开发  
    </a>  
</li>
```

示例中，如果要用文本获取`'前沿技术'`的`<a>`元素，可以这样写：

```
ele = tab.ele('text:前沿技术')  
# 或  
ele = tab.ele('@text():前沿技术')
```

这两种写法都能获取到包含直接文本的元素。

但如果要用文本获取`<li>`元素，就获取不到，因为文本不是`<li>`的直接内容。

我们可以这样写：

```
ele = tab.ele('tag:li@@text():前沿技术')
```

`@@text()`与`@text()`不同之处在于，前者可以搜索整个元素内所有文本，而不仅仅是直接文本，因此能实现一些非常灵活的查找。

注意

需要注意的是，使用`@@`或`@|`时，`text()`不要作为唯一的查询条件，否则会定位到整个文档最高层的元素。

❌ 错误做法：

```
ele = tab.ele('@@text():前沿技术')  
ele = tab.ele('@|text():前沿技术@|text():程序开发')
```

⭕ 正确做法：

```
ele = tab.ele('tag:li@|text():前沿技术@|text():程序开发')
```

---

[上一页

🔦 概述](/browser_control/get_elements/intro)[下一页

🔦 页面或元素内查找](/browser_control/get_elements/find_in_object)

* [✅️️ 基本概念](#️️-基本概念)
  + [📌 简单示例](#-简单示例)
* [✅️️ 基本逻辑](#️️-基本逻辑)
  + [📌 单属性匹配符 `@`](#-单属性匹配符-)
  + [📌 多属性与匹配符 `@@`](#-多属性与匹配符-)
  + [📌 多属性或匹配符 `@|`](#-多属性或匹配符-)
  + [📌 否定匹配符 `@!`](#-否定匹配符-)
  + [📌 混合使用](#-混合使用)
* [✅️️ 匹配模式](#️️-匹配模式)
  + [📌 精确匹配 `=`](#-精确匹配-)
  + [📌 模糊匹配 `:`](#-模糊匹配-)
  + [📌 匹配开头 `^`](#-匹配开头-)
  + [📌 匹配结尾 `$`](#-匹配结尾-)
* [✅️️ 常用语法](#️️-常用语法)
  + [📌 id 匹配符 `#`](#-id-匹配符-)
  + [📌 class 匹配符 `.`](#-class-匹配符-)
  + [📌 文本匹配符 `text`](#-文本匹配符-text)
  + [📌 类型匹配符 `tag`](#-类型匹配符-tag)
  + [📌 css selector 匹配符 `css`](#-css-selector-匹配符-css)
  + [📌 xpath 匹配符 `xpath`](#-xpath-匹配符-xpath)
  + [📌 selenium 的 loc 元组](#-selenium-的-loc-元组)
* [✅️️ `@@text()`的技巧](#️️-text的技巧)

* 🚀 控制浏览器
* 🛰️ 获取网页信息

本页总览

🛰️ 获取网页信息
=========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

成功访问网页后，可使用 Tab 对象属性和方法获取页面信息。

✅️️ 页面信息[​](#️️-页面信息 "✅️️ 页面信息的直接链接")
-------------------------------------

### 📌 `html`[​](#-html "-html的直接链接")

此属性返回当前页面 html 文本。

注意

html 文本不包含`<iframe>`元素内容。

**返回类型：**`str`

---

### 📌 `json`[​](#-json "-json的直接链接")

此属性把请求内容解析成 json。

假如用浏览器访问会返回 `*.json` 文件的 url，浏览器会把 json 数据显示出来，这个参数可以把这些数据转换为`dict`格式。

**返回类型：**`dict`

---

### 📌 `title`[​](#-title "-title的直接链接")

此属性返回当前页面`title`文本。

**返回类型：**`str`

---

### 📌 `user_agent`[​](#-user_agent "-user_agent的直接链接")

此属性返回当前页面 user agent 信息。

**返回类型：**`str`

---

### 📌 `save()`[​](#-save "-save的直接链接")

把当前页面保存为文件，同时返回保存的内容。

如果`path`和`name`参数都为`None`，只返回内容，不保存文件。

Page 对象和 Tab 对象有这个方法。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `path` | `str` `Path` | `None` | 保存路径，为`None`且`name`不为`None`时保存到当前路径 |
| `name` | `str` | `None` | 保存的文件名，为`None`且`path`不为`None`时使用 title 值 |
| `as_pdf` | `bool` | `False` | 为`Ture`保存为 pdf，否则保存为 mhtml 且忽略`kwargs`参数 |
| `**kwargs` | 多种 | 无 | pdf 生成参数 |

pdf 生成参数包括：`landscape`, `displayHeaderFooter`, `printBackground`, `scale`, `paperWidth`, `paperHeight`, `marginTop`, `marginBottom`, `marginLeft`, `marginRight`, `pageRanges`, `headerTemplate`, `footerTemplate`, `preferCSSPageSize`, `generateTaggedPDF`, `generateDocumentOutline`

| 返回类型 | 说明 |
| --- | --- |
| `str` | `as_pdf`为`False`时返回 mhtml 文本 |
| `bytes` | `as_pdf`为`True`时返回文件字节数据 |

---

✅️️ 运行状态信息[​](#️️-运行状态信息 "✅️️ 运行状态信息的直接链接")
-------------------------------------------

### 📌 `url`[​](#-url "-url的直接链接")

此属性返回当前访问的 url。

**返回类型：**`str`

---

### 📌 `tab_id`[​](#-tab_id "-tab_id的直接链接")

**返回类型：**`str`

此属性返回当前标签页的 id。

---

### 📌 `states.is_loading`[​](#-statesis_loading "-statesis_loading的直接链接")

此属性返回页面是否正在加载状态。

**返回类型：**`bool`

---

### 📌 `states.is_alive`[​](#-statesis_alive "-statesis_alive的直接链接")

此属性返回页面是否仍然可用，标签页已关闭则返回`False`。

**返回类型：**`bool`

---

### 📌 `states.ready_state`[​](#-statesready_state "-statesready_state的直接链接")

此属性返回页面当前加载状态，有 4 种：

* `'connecting'`： 网页连接中
* `'loading'`：表示文档还在加载中
* `'interactive'`：DOM 已加载，但资源未加载完成
* `'complete'`：所有内容已完成加载

**返回类型：**`str`

---

### 📌 `url_available`[​](#-url_available "-url_available的直接链接")

此属性以布尔值返回当前链接是否可用。

**返回类型：**`bool`

---

### 📌 `states.has_alert`[​](#-stateshas_alert "-stateshas_alert的直接链接")

此属性以布尔值返回页面是否存在弹出框。

**返回类型：**`bool`

---

✅️️ 窗口信息[​](#️️-窗口信息 "✅️️ 窗口信息的直接链接")
-------------------------------------

### 📌 `rect.size`[​](#-rectsize "-rectsize的直接链接")

以`tuple`返回页面大小，格式：(宽, 高)。

**返回类型：**`Tuple[int, int]`

---

### 📌 `rect.window_size`[​](#-rectwindow_size "-rectwindow_size的直接链接")

此属性以`tuple`返回窗口大小，格式：(宽, 高)。

**返回类型：**`Tuple[int, int]`

---

### 📌 `rect.window_location`[​](#-rectwindow_location "-rectwindow_location的直接链接")

此属性以`tuple`返回窗口在屏幕上的坐标，左上角为(0, 0)。

**返回类型：**`Tuple[int, int]`

---

### 📌 `rect.window_state`[​](#-rectwindow_state "-rectwindow_state的直接链接")

此属性以返回窗口当前状态，有`'normal'`、`'fullscreen'`、`'maximized'`、 `'minimized'`几种。

**返回类型：**`str`

---

### 📌 `rect.viewport_size`[​](#-rectviewport_size "-rectviewport_size的直接链接")

此属性以`tuple`返回视口大小，不含滚动条，格式：(宽, 高)。

**返回类型：**`Tuple[int, int]`

---

### 📌 `rect.viewport_size_with_scrollbar`[​](#-rectviewport_size_with_scrollbar "-rectviewport_size_with_scrollbar的直接链接")

此属性以`tuple`返回浏览器窗口大小，含滚动条，格式：(宽, 高)。

**返回类型：**`Tuple[int, int]`

---

### 📌 `rect.page_location`[​](#-rectpage_location "-rectpage_location的直接链接")

此属性以`tuple`返回页面左上角在屏幕中坐标，左上角为(0, 0)。

**返回类型：**`Tuple[int, int]`

---

### 📌 `rect.viewport_location`[​](#-rectviewport_location "-rectviewport_location的直接链接")

此属性以`tuple`返回视口在屏幕中坐标，左上角为(0, 0)。

**返回类型：**`Tuple[int, int]`

---

### 📌 `rect.scroll_position`[​](#-rectscroll_position "-rectscroll_position的直接链接")

此属性返回页面滚动条位置，格式：(x, y)。

**类型：**`Tuple[float, float]`

---

  

✅️️ 配置参数信息[​](#️️-配置参数信息 "✅️️ 配置参数信息的直接链接")
-------------------------------------------

### 📌 `timeout`[​](#-timeout "-timeout的直接链接")

此属性为整体默认超时时间（秒），包括元素查找、点击、处理提示框、列表选择等需要用到超时设置的地方，都以这个数据为默认值。默认为`10`。

**返回类型：**`int`、`float`

---

### 📌 `timeouts`[​](#-timeouts "-timeouts的直接链接")

此属性以字典方式返回三种超时时间（秒）。

* `'base'`：与`timeout`属性是同一个值
* `'page_load'`：用于等待页面加载
* `'script'`：用于等待脚本执行

**返回类型：**`dict`

```
print(tab.timeouts)
```

**输出：**

```
{'base': 10, 'page_load': 30.0, 'script': 30.0}
```

---

### 📌 `retry_times`[​](#-retry_times "-retry_times的直接链接")

此属性为网络连接失败时的重试次数，默认为`3`。

**返回类型：**`int`

---

### 📌 `retry_interval`[​](#-retry_interval "-retry_interval的直接链接")

此属性为网络连接失败时的重试等待间隔秒数，默认为`2`。

**返回类型：**`int`、`float`

---

### 📌 `load_mode`[​](#-load_mode "-load_mode的直接链接")

此属性返回页面加载策略，有 3 种：

* `'normal'`：等待页面所有资源完成加载
* `'eager'`：DOM 加载完成即停止
* `'none'`：页面完成连接即停止

**返回类型：**`str`

---

✅️️ cookies 和缓存信息[​](#️️-cookies-和缓存信息 "✅️️ cookies 和缓存信息的直接链接")
----------------------------------------------------------------

### 📌 `cookies()`[​](#-cookies "-cookies的直接链接")

此方法以列表方式返回 cookies 信息。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `all_domains` | `bool` | `False` | 是否返回所有 cookies，为`False`只返回当前 url 的 |
| `all_info` | `bool` | `False` | 返回的 cookies 是否包含所有信息，`False`时只包含`name`、`value`、`domain`信息 |

| 返回类型 | 说明 |
| --- | --- |
| `CookiesList` | cookies 组成的列表 |

**示例：**

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.get('https://www.baidu.com')  
  
for i in tab.cookies():  
    print(i)
```

**输出：**

```
{'domain': '.baidu.com', 'domain_specified': True, ......}  
......
```

---

### 📌 指定返回类型[​](#-指定返回类型 "📌 指定返回类型的直接链接")

`cookies()`方法返回的列表可转换为其它指定格式。

* `cookies().as_str()`：`'name1=value1; name2=value2'`格式的字符串
* `cookies().as_dict()`：`{name1: value1, name2: value2}`格式的字典
* `cookies().as_json()`：json 格式的字符串

说明

`as_str()`和`as_dict()`都只会保留`'name'`和`'value'`字段。

---

### 📌 `session_storage()`[​](#-session_storage "-session_storage的直接链接")

此方法用于获取 sessionStorage 信息，可获取全部或单个项。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `item` | `str` | `None` | 要获取的项目，为`None`则返回全部项目组成的字典 |

| 返回类型 | 说明 |
| --- | --- |
| `dict` | `item`参数为`None`时返回所有项目 |
| `str` | 指定`item`时返回该项目内容 |

---

### 📌 `local_storage()`[​](#-local_storage "-local_storage的直接链接")

此方法用于获取 localStorage 信息，可获取全部或单个项。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `item` | `str` | `None` | 要获取的项目，为`None`则返回全部项目组成的字典 |

| 返回类型 | 说明 |
| --- | --- |
| `dict` | `item`参数为`None`时返回所有项目 |
| `str` | 指定`item`时返回该项目内容 |

---

[上一页

🛰️ 页面交互](/browser_control/page_operation)[下一页

🔦 概述](/browser_control/get_elements/intro)

* [✅️️ 页面信息](#️️-页面信息)
  + [📌 `html`](#-html)
  + [📌 `json`](#-json)
  + [📌 `title`](#-title)
  + [📌 `user_agent`](#-user_agent)
  + [📌 `save()`](#-save)
* [✅️️ 运行状态信息](#️️-运行状态信息)
  + [📌 `url`](#-url)
  + [📌 `tab_id`](#-tab_id)
  + [📌 `states.is_loading`](#-statesis_loading)
  + [📌 `states.is_alive`](#-statesis_alive)
  + [📌 `states.ready_state`](#-statesready_state)
  + [📌 `url_available`](#-url_available)
  + [📌 `states.has_alert`](#-stateshas_alert)
* [✅️️ 窗口信息](#️️-窗口信息)
  + [📌 `rect.size`](#-rectsize)
  + [📌 `rect.window_size`](#-rectwindow_size)
  + [📌 `rect.window_location`](#-rectwindow_location)
  + [📌 `rect.window_state`](#-rectwindow_state)
  + [📌 `rect.viewport_size`](#-rectviewport_size)
  + [📌 `rect.viewport_size_with_scrollbar`](#-rectviewport_size_with_scrollbar)
  + [📌 `rect.page_location`](#-rectpage_location)
  + [📌 `rect.viewport_location`](#-rectviewport_location)
  + [📌 `rect.scroll_position`](#-rectscroll_position)
* [✅️️ 配置参数信息](#️️-配置参数信息)
  + [📌 `timeout`](#-timeout)
  + [📌 `timeouts`](#-timeouts)
  + [📌 `retry_times`](#-retry_times)
  + [📌 `retry_interval`](#-retry_interval)
  + [📌 `load_mode`](#-load_mode)
* [✅️️ cookies 和缓存信息](#️️-cookies-和缓存信息)
  + [📌 `cookies()`](#-cookies)
  + [📌 指定返回类型](#-指定返回类型)
  + [📌 `session_storage()`](#-session_storage)
  + [📌 `local_storage()`](#-local_storage)

* 🚀 控制浏览器
* 🛰️ iframe 操作

本页总览

🛰️ iframe 操作
============

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

`<iframe>`元素是一种特殊的元素，它既是元素，也是页面。

DrissionPage 无需切入切出即可处理`<iframe>`元素。
可实现跨级元素查找、元素内部单独跳转、同时操作`<iframe>`内外元素、多线程控制多个`<iframe>`等操作。

✅️ 获取`<iframe>`对象[​](#️-获取iframe对象 "️-获取iframe对象的直接链接")
-------------------------------------------------------

获取`<iframe>`对象的方法有两种，可用获取普通元素的方式获取，或者用`get_frame()`方法获取。

推荐优先使用`get_frame()`方法，因为当作普通元素获取时，IDE 无法正确识别获取到的是`<iframe>`元素。

### 📌 `get_frame()`[​](#-get_frame "-get_frame的直接链接")

此方法用于获取页面中一个`<frame>`或`<iframe>`对象。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `loc_ind_ele` | `str` `int` `ChromiumFrame` | 必填 | 定位符 `<iframe>`元素序号（从`1`开始，负数表示倒数） `ChromiumFrame对象` `id`属性内容 `name`属性内容 |
| `timeout` | `float` | `None` | 超时时间（秒），为`None`时使用页面超时时间 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumFrame` | `<frame>`或`<iframe>`元素对象 |
| `NoneElement` | 找不到时返回`NoneElement` |

注意

需要特别注意的是，如果页面中有嵌套的`<iframe>`，用序号获取的方式会存在不准确。
比如，用`get_frames()`可获取到 6 个元素，但用`get_frame(6)`却获取不到最后一个。
这是因为有两个`<iframe>`是嵌套关系，导致获取不准确。

**示例：**

```
# 使用定位符获取  
iframe = tab.get_frame('t:iframe')  
  
# 获取第1个iframe  
iframe = tab.get_frame(1)
```

---

### 📌 `get_frames()`[​](#-get_frames "-get_frames的直接链接")

此方法用于获取页面中多个符合条件的`<frame>`或`<iframe>`对象。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` | `None` | 定位符，为`None`时返回所有 |
| `timeout` | `float` | `None` | 超时时间（秒），为`None`时使用页面超时时间 |

| 返回类型 | 说明 |
| --- | --- |
| `List[ChromiumFrame]` | `<frame>`或`<iframe>`元素对象组成的列表 |

提醒

获取所有`<iframe>`会很慢，而且浪费资源，一般使用获取需要用到的就好。

---

### 📌 普通元素方式[​](#-普通元素方式 "📌 普通元素方式的直接链接")

可以用获取普通元素的方式获取`<iframe>`对象：

```
iframe = tab('t:iframe')
```

这个`ChromiumFrame`对象既是页面也是元素。由于 IDE 不会提示`<iframe>`
元素对象相关的属性和方法，因此用这种方式获取时建议再用`get_frame()`包装一下：

```
iframe = tab('t:iframe')  
iframe = tab.get_frame(iframe)
```

---

✅️ 查找`<iframe>`内元素[​](#️-查找iframe内元素 "️-查找iframe内元素的直接链接")
----------------------------------------------------------

当`<iframe>`与标签页是同域的，我们并不需要先切入`<iframe>`，就可以获取到里面的元素。

如果是异域的，则先要获取这个标签页的`ChromiumFrame`对象，再用这个对象在自己内部搜索。

### 📌 页面跨`<iframe>`查找[​](#-页面跨iframe查找 "-页面跨iframe查找的直接链接")

如果`<iframe>`元素的网址和主页面是同域的，我们可以直接用页面对象查找`<iframe>`内部元素，而无需先获取`ChromiumFrame`对象。

以下示例页面中有一个`<iframe>`元素，和标签页是同域的，可直接通过 Tab 对象查找它内部的元素。

只要是同域名的，无论跨多少层`<iframe>`都能用页面对象直接获取。

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.get('http://DrissionPage.cn/demos/iframe_same_domain.html')  
ele = tab('概述')  
print(ele)
```

**输出：**

```
<ChromiumElement h2 class='anchor anchorWithStickyNavbar_LWe7' id='️-概述'>
```

---

### 📌 在`<iframe>`内查找[​](#-在iframe内查找 "-在iframe内查找的直接链接")

如果`<iframe>`跟当前标签页是不同域名的，不能使用页面对象直接查找其中元素，只能先获取其`ChromiumFrame`元素对象，再在这个对象中查找。

即使是同域的，也可以通过这种方法查找。

但创建`ChromiumFrame`对象会增加系统资源的使用，一般建议异域的才创建对象。

以下示例页面中有一个`<iframe>`元素，和标签页是不同域的，需要先获取`ChromiumFrame`对象，再在里面找元素。

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.get('http://DrissionPage.cn/demos/iframe_diff_domain.html')  
iframe = tab.get_frame('t:iframe')  
ele = iframe('网易首页')  
print(ele)
```

**输出：**

```
<ChromiumElement a class='ntes-nav-index-title ntes-nav-entry-wide c-fl' href='https://www.163.com/' title='网易首页'>
```

---

  

✅️ 方法和属性[​](#️-方法和属性 "✅️ 方法和属性的直接链接")
-------------------------------------

正如上面所说，`ChromiumFrame`既是元素也是页面，它可以获取自身元素方面的属性或执行操作。

详见相关章节。

```
iframe.tag  
iframe.html  
iframe.remove_attr()  
iframe.states.is_alive  
iframe.get()  
iframe.get_screenshot()  
# 等等
```

[上一页

🛰️ 获取元素信息](/browser_control/get_ele_info)[下一页

🛰️ 动作链](/browser_control/actions)

* [✅️ 获取`<iframe>`对象](#️-获取iframe对象)
  + [📌 `get_frame()`](#-get_frame)
  + [📌 `get_frames()`](#-get_frames)
  + [📌 普通元素方式](#-普通元素方式)
* [✅️ 查找`<iframe>`内元素](#️-查找iframe内元素)
  + [📌 页面跨`<iframe>`查找](#-页面跨iframe查找)
  + [📌 在`<iframe>`内查找](#-在iframe内查找)
* [✅️ 方法和属性](#️-方法和属性)

* 🚀 控制浏览器
* 🛰️ 概述

本页总览

🛰️ 概述
=====

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

✅️️ 基本逻辑[​](#️️-基本逻辑 "✅️️ 基本逻辑的直接链接")
-------------------------------------

操作浏览器的基本逻辑如下：

1. 创建浏览器对象，用于启动或接管浏览器
2. 获取一个 Tab 对象
3. 使用 Tab 对象访问网址
4. 使用 Tab 对象获取标签页内需要的元素对象
5. 使用元素对象进行交互

除此以外，还能执行更为复杂的操作，如执行 js 代码、监听网络数据、下载文件等。这些在后面的章节再介绍。

**示例：** 在百度搜索 “Drissionpage”，并打印结果。

```
# 导入  
from DrissionPage import Chromium  
  
# 连接浏览器  
browser = Chromium()    
# 获取标签页对象  
tab = browser.latest_tab    
# 访问网页  
tab.get('https://www.baidu.com')    
# 获取文本框元素对象  
ele = tab.ele('#kw')  
# 向文本框元素对象输入文本  
ele.input('DrissionPage')    
# 点击按钮，上两行的代码可以缩写成这样  
tab('#su').click()    
# 获取所有<h3>元素  
links = tab.eles('tag:h3')    
# 遍历并打印结果  
for link in links:    
    print(link.text)
```

---

✅️️ 浏览器对象[​](#️️-浏览器对象 "✅️️ 浏览器对象的直接链接")
----------------------------------------

即`Chromium`对象，用于管理浏览器整体相关的操作。

如标签页管理、获取浏览器信息、设置整体运行参数等。

```
from DrissionPage import Chromium  
  
browser = Chromium()  # 创建浏览器对象  
browser.set.retry_times(10)  # 设置整体运行参数  
tab = browser.latest_tab  # 获取Tab对象  
browser.quit()  # 关闭浏览器
```

---

✅️️ 标签页对象[​](#️️-标签页对象 "✅️️ 标签页对象的直接链接")
----------------------------------------

Tab 对象从浏览器对象获取，每个 Tab 对象对应浏览器上一个实际的标签页。

大部分操作都使用 Tab 对象进行，如访问网站、调整窗口大小、监听网络等。

默认情况下每个标签页只有一个 Tab 对象，关闭单例模式后可用多个 Tab 对象同时控制一个标签页。

```
from DrissionPage import Chromium  
  
browser = Chromium()  
tab1 = browser.latest_tab  # 获取最后激活的标签页对象  
tab1.get('http://DrissionPage.cn')  # 标签页访问一个网址  
tab2 = browser.new_tab('https://www.baidu.com')  # 新建一个标签页并访问网址  
tab3 = browser.get_tab(title='DrissionPage')  # 按条件获取标签页对象
```

---

  

✅️️ 元素对象[​](#️️-元素对象 "✅️️ 元素对象的直接链接")
-------------------------------------

元素对象`ChromiumElemet`是交互的执行者，如点击、文本输入、获取元素信息等。

元素对象可从 Tab 对象获取，也可从另一个元素对象通过内部查找或相对定位的方式获取。

### 📌 对象内部查找[​](#-对象内部查找 "📌 对象内部查找的直接链接")

Tab 对象和 元素对象都有`ele()`方法，用于在其内部查找指定元素。

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.get('http://DrissionPage.cn')  
ele = tab.ele('text=文档')  # 获取文本为“文档”的元素  
ele.click()  # 点击该元素
```

---

### 📌 相对位置查找[​](#-相对位置查找 "📌 相对位置查找的直接链接")

可先获取一个元素对象，然后以这个元素为基准定位其内部或指定相对关系的元素。

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.get('http://DrissionPage.cn')  
ele1 = tab.ele('text=文档')  # 获取文本为“文档”的元素  
ele2 = ele1.next()  # 获取ele1的后一个元素  
ele2.click()  # 点击该元素
```

---

[下一页

🛰️ 连接浏览器](/browser_control/connect_browser)

* [✅️️ 基本逻辑](#️️-基本逻辑)
* [✅️️ 浏览器对象](#️️-浏览器对象)
* [✅️️ 标签页对象](#️️-标签页对象)
* [✅️️ 元素对象](#️️-元素对象)
  + [📌 对象内部查找](#-对象内部查找)
  + [📌 相对位置查找](#-相对位置查找)

* 🚀 控制浏览器
* 🛰️ 监听网络数据

本页总览

🛰️ 监听网络数据
=========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

许多网页的数据来自接口，在网站使用过程中动态加载，如使用 JS 加载内容的翻页列表。

这些数据通常以 json 形式发送，浏览器接收后，对其进行解析，再加载到 DOM 相应位置。

做数据采集的时候，我们往往从 DOM 中去获取解析后数据的，可能存在数据不全、加载响应不及时、难以判断加载完成等问题。

如果我们可以拿到浏览器收发的数据包，根据数据包状态判断下一步操作，甚至直接获取数据，岂不是爽爆了？

DrissionPage 每个页面对象（包括 Tab 和 Frame 对象）内置了一个监听器，专门用于抓取浏览器数据包。

可以提供等待和捕获指定数据包，实时返回指定数据包功能。

✅️ 示例[​](#️-示例 "✅️ 示例的直接链接")
----------------------------

先看两个示例了解监听器工作方式。

注意

要先启动监听，再执行动作，`listen.start()`之前的数据包是获取不到的。

### 📌 等待并获取[​](#-等待并获取 "📌 等待并获取的直接链接")

点击下一页，等待数据包，再点击下一页，循环。

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.get('https://gitee.com/explore/all')  # 访问网址，这行产生的数据包不监听  
  
tab.listen.start('gitee.com/explore')  # 开始监听，指定获取包含该文本的数据包  
for _ in range(5):  
    tab('@rel=next').click()  # 点击下一页  
    res = tab.listen.wait()  # 等待并获取一个数据包  
    print(res.url)  # 打印数据包url
```

**输出：**

```
https://gitee.com/explore/all?page=2  
https://gitee.com/explore/all?page=3  
https://gitee.com/explore/all?page=4  
https://gitee.com/explore/all?page=5  
https://gitee.com/explore/all?page=6
```

---

### 📌 实时获取[​](#-实时获取 "📌 实时获取的直接链接")

跟上一个示例做同样的事情，不过换一种方式。

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.listen.start('gitee.com/explore')  # 开始监听，指定获取包含该文本的数据包  
tab.get('https://gitee.com/explore/all')  # 访问网址  
  
i = 0  
for packet in tab.listen.steps():  
    print(packet.url)  # 打印数据包url  
    tab('@rel=next').click()  # 点击下一页  
    i += 1  
    if i == 5:  
        break
```

---

✅️ 设置目标和启动监听[​](#️-设置目标和启动监听 "✅️ 设置目标和启动监听的直接链接")
-------------------------------------------------

### 📌 `listen.start()`[​](#-listenstart "-listenstart的直接链接")

此方法用于启动监听器，启动同时可设置获取的目标特征。

可设置多个特征，符合条件的数据包会被获取。

如果监听未停止时调用这个方法，可清除已抓取的队列。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `targets` | `str` `list` `tuple` `set` | `None` | 要匹配的数据包 url 特征，可用列表指定多个，为`True`时获取所有 |
| `is_regex` | `bool` | `None` | 设置的 target 是否正则表达式，为`None`时保持原来设置 |
| `method` | `str` `list` `tuple` `set` | `None` | 设置监听的请求类型，可指定多个，默认`('GET', 'POST')`，为`True`时监听所有，为`None`时保持原来设置 |
| `res_type` | `str` `list` `tuple` `set` | `None` | 设置监听的 ResourceType 类型，可指定多个，为`True`时监听所有，为`None`时保持原来设置 |

**返回：** `None`

注意

当`targets`不为`None`，`is_regex`会自动设为`False`。  
即如要使用正则，每次设置`targets`时需显式指定`is_regex=True`。

---

### 📌 `listen.set_targets()`[​](#-listenset_targets "-listenset_targets的直接链接")

此方法可在监听过程中修改监听目标，也可在监  听开始前设置。

如监听未启动，不会启动监听。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `targets` | `str` `list` `tuple` `set` | `True` | 要匹配的数据包 url 特征，可用列表指定多个，为`True`时获取所有 |
| `is_regex` | `bool` | `False` | 设置的 target 是否正则表达式 |
| `method` | `str` `list` `tuple` `set` | `('GET', 'POST')` | 设置监听的请求类型，可指定多个，默认`('GET', 'POST')`，为`True`时监听所有 |
| `res_type` | `str` `list` `tuple` `set` | `True` | 设置监听的 ResourceType 类型，可指定多个，为`True`时监听所有 |

**返回：** `None`

---

✅️ 等待和获取数据包[​](#️-等待和获取数据包 "✅️ 等待和获取数据包的直接链接")
----------------------------------------------

### 📌 `listen.wait()`[​](#-listenwait "-listenwait的直接链接")

此方法用于等待符合要求的数据包到达指定数量。

所有符合条件的数据包都会存储到队列，`wait()`实际上是逐个从队列中取结果，不用担心页面已刷走而丢包。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `count` | `int` | `1` | 需要捕捉的数据包数量 |
| `timeout` | `float` `None` | `None` | 超时时间（秒），为`None`无限等待 |
| `fit_count` | `bool` | `True` | 是否必需满足总数要求，如超时，为`True`返回`False`，为`False`返回已捕捉到的数据包 |
| `raise_err` | `bool` | `None` | 超时时是否抛出错误，为`None`时根据`Settings`设置，如不抛出，超时返回`False` |

| 返回类型 | 说明 |
| --- | --- |
| `DataPacket` | `count`为`1`且未超时，返回一个数据包对象 |
| `List[DataPacket]` | `count`大于`1`，未超时或`fit_count`为`False`，返回数据包对象组成的列表 |
| `False` | 超时且`fit_count`为`True`时 |

---

### 📌 `listen.steps()`[​](#-listensteps "-listensteps的直接链接")

此方法返回一个可迭代对象，用于`for`循环，每次循环可从中获取到的数据包。

可实现实时获取并返回数据包。

如果`timeout`超时，会中断循环。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `count` | `int` | `None` | 需捕获的数据包总数，为`None`表示无限 |
| `timeout` | `float` `None` | `None` | 每个数据包等待时间（秒），为`None`表示无限等待 |
| `gap` | `int` | `1` | 每接收到多少个数据包返回一次数据 |

| 返回类型 | 说明 |
| --- | --- |
| `DataPacket` | `gap`为`1`时，返回一个数据包对象 |
| `List[DataPacket]` | `gap`大于`1`，返回数据包对象组成的列表 |

---

### 📌 `listen.wait_silent()`[​](#-listenwait_silent "-listenwait_silent的直接链接")

此方法用于等待所有指定的请求完成。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` `None` | `None` | 等待时间（秒），为`None`表示无限等待 |
| `targets_only` | `bool` | `False` | 是否只等待`targets`指定的请求结束 |
| `limit` | `int` | `0` | 剩下多少个连接时视为结束 |

| 返回类型 | 说明 |
| --- | --- |
| `bool` | 是否等待成功 |

---

  

✅️ 暂停和恢复[​](#️-暂停和恢复 "✅️ 暂停和恢复的直接链接")
-------------------------------------

### 📌 `listen.pause()`[​](#-listenpause "-listenpause的直接链接")

此方法用于暂停监听。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `clear` | `bool` | `True` | 是否清空已获取队列 |

**返回：** `None`

---

### 📌 `listen.resume()`[​](#-listenresume "-listenresume的直接链接")

此方法用于继续暂停的监听。

**参数：** 无

**返回：**`None`

---

### 📌 `listen.stop()`[​](#-listenstop "-listenstop的直接链接")

此方法用于终止监听器的运行，会清空已获取的队列，不清空 targets。

**参数：** 无

**返回：**`None`

---

✅️ `DataPacket`对象[​](#️-datapacket对象 "️-datapacket对象的直接链接")
-----------------------------------------------------------

`DataPacket`对象是获取到的数据包结果对象，包含了数据包各种信息。

### 📌 `对象属性`[​](#-对象属性 "-对象属性的直接链接")

| 属性名称 | 数据类型 | 说明 |
| --- | --- | --- |
| `tab_id` | `str` | 产生这个请求的标签页的 id |
| `frameId` | `str` | 产生这个请求的框架 id |
| `target` | `str` | 产生这个请求的监听目标 |
| `url` | `str` | 数据包请求网址 |
| `method` | `str` | 请求类型 |
| `is_failed` | `bool` | 是否连接失败 |
| `resourceType` | `str` | 资源类型 |
| `request` | `Request` | 保存请求信息的对象 |
| `response` | `Response` | 保存响应信息的对象 |
| `fail_info` | `FailInof` | 保存连接失败信息的对象 |

### 📌 `wait_extra_info()`[​](#-wait_extra_info "-wait_extra_info的直接链接")

有些数据包有`extra_info`数据，但这些数据可能会迟于数据包返回，用这个方法可以等待这些数据加载到数据包对象。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` `None` | `None` | 超时时间（秒），`None`为无限等待 |

| 返回类型 | 说明 |
| --- | --- |
| `bool` | 是否等待成功 |

### 📌 `Request`对象[​](#-request对象 "-request对象的直接链接")

`Request`对象是`DataPacket`对象内用于保存请求信息的对象，有以下属性：

| 属性名称 | 数据类型 | 说明 |
| --- | --- | --- |
| `url` | `str` | 请求的网址 |
| `method` | `str` | 请求类型 |
| `params` | `dict` | 以`dict`格式返回 url 中的参数 |
| `headers` | `CaseInsensitiveDict` | 以大小写不敏感字典返回 headers 数据 |
| `cookies` | `List[dict]` | 返回发送的 cookies |
| `postData` | `str` `dict` | post 类型的请求所提交的数据，json 以`dict`格式返回 |

除以上常用属性，还有以 下属性，自行体会：

urlFragment、hasPostData、postDataEntries、mixedContentType、initialPriority、referrerPolicy、isLinkPreload、trustTokenParams、isSameSite

---

### 📌 `Response`对象[​](#-response对象 "-response对象的直接链接")

`Response`对象是`DataPacket`对象内用于保存响应信息的对象，有以下属性：

| 属性名称 | 数据类型 | 说明 |
| --- | --- | --- |
| `url` | `str` | 请求的网址 |
| `headers` | `CaseInsensitiveDict` | 以大小写不敏感字典返回 headers 数据 |
| `body` | `str` `bytes` `dict` | 如果是 json 格式，转换为`dict`；如果是 base64 格式，转换为`bytes`，其它格式直接返回文本 |
| `raw_body` | `str` | 未被处理的 body 文本 |
| `status` | `int` | 请求状态 |
| `statusText` | `str` | 请求状态文本 |

除以上属性，还有以下属性，自行体会：

headersText、mimeType、requestHeaders、requestHeadersText、connectionReused、connectionId、remoteIPAddress、remotePort、fromDiskCache、fromServiceWorker、fromPrefetchCache、encodedDataLength  、timing、serviceWorkerResponseSource、responseTime、cacheStorageCacheName、protocol、alternateProtocolUsage、securityState、securityDetails

---

### 📌 `FailInfo`对象[​](#-failinfo对象 "-failinfo对象的直接链接")

`FailInfo`对象是`DataPacket`对象内用于保存连接失败信息的对象，有以下属性：

| 属性名称 | 数据类型 | 说明 |
| --- | --- | --- |
| `errorText` | `str` | 错误信息文本 |
| `canceled` | `bool` | 是否取消 |
| `blockedReason` | `str` | 拦截原因 |
| `corsErrorStatus` | `str` | cors 错误状态 |

[上一页

🛰️ 等待](/browser_control/waiting)[下一页

🛰️ 获取控制台信息](/browser_control/console)

* [✅️ 示例](#️-示例)
  + [📌 等待并获取](#-等待并获取)
  + [📌 实时获取](#-实时获取)
* [✅️ 设置目标和启动监听](#️-设置目标和启动监听)
  + [📌 `listen.start()`](#-listenstart)
  + [📌 `listen.set_targets()`](#-listenset_targets)
* [✅️ 等待和获取数据包](#️-等待和获取数据包)
  + [📌 `listen.wait()`](#-listenwait)
  + [📌 `listen.steps()`](#-listensteps)
  + [📌 `listen.wait_silent()`](#-listenwait_silent)
* [✅️ 暂停和恢复](#️-暂停和恢复)
  + [📌 `listen.pause()`](#-listenpause)
  + [📌 `listen.resume()`](#-listenresume)
  + [📌 `listen.stop()`](#-listenstop)
* [✅️ `DataPacket`对象](#️-datapacket对象)
  + [📌 `对象属性`](#-对象属性)
  + [📌 `wait_extra_info()`](#-wait_extra_info)
  + [📌 `Request`对象](#-request对象)
  + [📌 `Response`对象](#-response对象)
  + [📌 `FailInfo`对象](#-failinfo对象)

* 🚀 控制浏览器
* 🛰️ 模式切换

本页总览

🛰️ 模式切换
=======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

`MixTab`和`WebPage`有两种模式，d 模式用于控制浏览器，s 模式使用`requests`收发数据包。

两种模式访问页面和提取数据的逻辑是一致的，使用同一套 api。

每个标签页对象创建时都处于 d 模式。

使用`change_mode()`方法进行切换。模式切换的时候会同步登录信息。

s 模式下仍然可以控制浏览器，但因为共用 api，`ele()`等两种模式共用的方法，查找对象是`requests`的结果，而非浏览器。

因此 s 模式下要控制浏览器，只能调用 d 模式独有的功能。

在切换模式前已获取的元素对象则可继续操作。

Tips

切换到 s 模式后，如不再需要浏览器，可以用`close()`或`quit()`方法关闭标签页或浏览器。标签页对象继续用于收发数据包。

✅️️ 示例[​](#️️-示例 "✅️️ 示例的直接链接")
-------------------------------

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.get('http://DrissionPage.cn')  
print(tab.title)  # 打印d模式下网页title  
tab.change_mode()  # 切换到s模式，切换时会自动访问d模式的url  
print(tab.title)  # 打印s模式下网页title
```

**输出：**

```
DrissionPage官网  
DrissionPage官网
```

---

✅️️ 相关属性和方法[​](#️️-相关属性和方法 "✅️️ 相关属性和方法的直接链接")
----------------------------------------------

### 📌️ `mode`[​](#️-mode "️-mode的直接链接")

此属性返回当前模式。`'d'`或`'s'`。

**类型：**`str`

---

### 📌 `change_mode()`[​](#-change_mode "-change_mode的直接链接")

此方法用于切换运行模式。

切换模式时默认复制当前 cookies 到目标模式，且使用当前 url 进行跳转。

注意

切换模式时只同步 cookies，不同步 headers，如果网站要求特定的 headers 才能访问，就会卡住直到超时。
这时可以设置`go`为`False`，切换 s 模式后再自己构造 headers 访问。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `mode` | `str` `None` | `None` | 接收`'s'`或`'d'`，以切换到指定模式 接收`None`则切换到与当前相对的另一个模式 |
| `go` | `bool` | `True` | 目标模式是否跳转到原模式的 url |
| `copy_cookies` | `bool` | `True` | 切换时是否复制 cookies 到目标模式 |

**返回：**`None`

---

### 📌 `cookies_to_session()`[​](#-cookies_to_session "-cookies_to_session的直接链接")

此方法用于复制浏览器当前页面的 cookies 到`Session`对象。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `copy_user_agent` | `bool` | `True` | 是否复制 user agent 信息 |

**返回：**`None`

### 📌 `cookies_to_browser()`[​](#-cookies_to_browser "-cookies_to_browser的直接链接")

此方法用于把`Session`对象的 cookies 复制到浏览器。

**参数：** 无

**返回：**`None`

---

✅️️ 说明[​](#️️-说明 "✅️️ 说明的直接链接")
-------------------------------

* 主要的 api 两种模式是共用的，如`get()`，d 模式下控制浏览跳转，s 模式下控制`Session`对象跳转
* s 模式下获取的元素对象为`SessionElement`，d 模式下为`ChromiumElement`等
* `post()`方法无论在哪种模式下都能使用
* s 模式下也能控制浏览器，但只能使用 d 模式独有功 能控制

---

[上一页

🛰️ 动作链](/browser_control/actions)[下一页

🛰️ 等待](/browser_control/waiting)

* [✅️️ 示例](#️️-示例)
* [✅️️ 相关属性和方法](#️️-相关属性和方法)
  + [📌️ `mode`](#️-mode)
  + [📌 `change_mode()`](#-change_mode)
  + [📌 `cookies_to_session()`](#-cookies_to_session)
  + [📌 `cookies_to_browser()`](#-cookies_to_browser)
* [✅️️ 说明](#️️-说明)

* 🚀 控制浏览器
* 🛰️ 页面交互

本页总览

🛰️ 页面交互
=======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

本节介绍浏览器页面交互功能。

一个 Tab 对象控制一个浏览器的标签页，是页面控制的主要单位。

✅️️ 页面跳转[​](#️️-页面跳转 "✅️️ 页面跳转的直接链接")
-------------------------------------

### 📌 `get()`[​](#-get "-get的直接链接")

详见 “访问网页” 章节。

---

### 📌 `back()`[​](#-back "-back的直接链接")

此方法用于在浏览历史中后退若干步。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `steps` | `int` | `1` | 后退步数 |

**返回：**`None`

**示例：**

```
tab.back(2)  # 后退两个网页
```

---

### 📌 `forward()`[​](#-forward "-forward的直接链接")

此方法用于在浏览历史中前进若干步。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `steps` | `int` | `1` | 前进步数 |

**返回：**`None`

```
tab.forward(2)  # 前进两步
```

---

### 📌 `refresh()`[​](#-refresh "-refresh的直接链接")

此方法用于刷新当前页面。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `ignore_cache` | `bool` | `False` | 刷新时是否忽略缓存 |

**返回：**`None`

**示例：**

```
tab.refresh()  # 刷新页面
```

---

### 📌 `stop_loading()`[​](#-stop_loading "-stop_loading的直接链接")

此方法用于强制停止当前页面加载。

**参数：** 无

**返回：**`None`

---

### 📌 `set.blocked_urls()`[​](#-setblocked_urls "-setblocked_urls的直接链接")

此方法用于设置忽略的连接。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `urls` | `str` `list` `tuple` `None` | 必填 | 要忽略的 url，可传入多个，可用`'*'`通配符，传入`None`时清空已设置的项 |

**返回：**`None`

**示例：**

```
tab.set.blocked_urls('*.css*')  # 设置不加载css文件
```

---

✅️️ 元素管理[​](#️️-元素管理 "✅️️ 元素管理的直接链接")
-------------------------------------

### 📌 `add_ele()`[​](#-add_ele "-add_ele的直接链接")

此方法用于创建一个元素。可选择是否插入到 DOM。

`html_or_info`传入元素完整 html 文本时，会插入到 DOM。如`insert_to`参数为`None`，插入到`body`元素。

传入元素信息（格式：`(tag, {name: value})`）时，如`insert_to`参数为`None`，不插入到 DOM。此时返回的元素需用 js 方式点击。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `html_or_info` | `str` `Tuple[str, dict]` | 必填 | 新元素的 html 文本或信息；为`tuple`可新建不加入到 DOM 的元素 |
| `insert_to` | `str` `ChromiumElement` `Tuple[str, str]` | `None` | 插入到哪个元素中，可接收元素对象和定位符；如为`None`，`html_or_info`是`str`时添加到 body，否则不添加到 DOM |
| `before` | `str` `ChromiumElement` `Tuple[str, str]` | `None` | 在哪个子节点前面插入，可接收对象和定位符，为`None`插入到父元素末尾 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 新建的元素对象 |

**添加一个可见的元素：**

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.get('https://www.baidu.com')  
html = '<a href="http://DrissionPage.cn" target="blank">DrissionPage </a> '  
ele = tab.add_ele(html, '#s-top-left', '新闻')  # 插入到导航栏  
ele.click()
```

**添加一个不可见的元素：**

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
info = ('a', {'innerText': 'DrissionPage', 'href': 'http://DrissionPage.cn', 'target': 'blank'})  
ele = tab.add_ele(info)  
ele.click('js')  # 需用js点击
```

---

### 📌 `remove_ele()`[​](#-remove_ele "-remove_ele的直接链接")

此方法用于从页面上删除一个元素。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `loc_or_ele` | `str` `Tuple[str, str]` `ChromiumElement` | 必填 | 要删除的元素，可以是元素或定位符 |

**返回：**`None`

**示例：**

```
# 删除一个已获得的元素  
ele = tab('tag:a')  
tab.remove_ele(ele)  
  
# 删除用定位符找到的元素  
tab.remove_ele('tag:a')
```

---

✅️️ 执行脚本或命令[​](#️️-执行脚本或命令 "✅️️ 执行脚本或命令的直接链接")
----------------------------------------------

### 📌 `run_js()`[​](#-run_js "-run_js的直接链接")

此方法用于执行 js 脚本。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `script` | `str` | 必填 | js 脚本文本或脚本文件路径 |
| `*args` | - | 无 | 传入的参数，按顺序在js文本中对应`arguments[0]`、`arguments[1]`... |
| `as_expr` | `bool` | `False` | 是否作为表达式运行，为`True`时`args`参数无效 |
| `timetout` | `float` | `None` | js 超时时间（秒），为`None`则使用页面`timeouts.script`设置 |

| 返回类型 | 说明 |
| --- | --- |
| `Any` | 脚本执行结果 |

**示例：**

```
# 用传入参数的方式执行 js 脚本显示弹出框显示 Hello world!  
tab.run_js('alert(arguments[0]+arguments[1]);', 'Hello', ' world!')
```

注意

* 如果`as_expr`为`True`，脚本应是返回一个结果的形式，并且不能有`return`
* 如果`as_expr`不为`True'，脚本应尽量写成一个方法。

---

### 📌 `run_js_loaded()`[​](#-run_js_loaded "-run_js_loaded的直接链接")

此方法用于运行 js 脚本，执行前等待页面加载完毕。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `script` | `str` | 必填 | js 脚本文本 |
| `*args` | - | 无 | 传入的参数，按顺序在js文本中对应`arguments[0]`、`arguments[1]`... |
| `as_expr` | `bool` | `False` | 是否作为表达式运行，为`True`时`args`参数无效 |
| `timetout` | `float` | `None` | js 超时时间（秒），为`None`则使用页面`timeouts.script`设置 |

| 返回类型 | 说明 |
| --- | --- |
| `Any` | 脚本执行结果 |

---

### 📌 `run_async_js()`[​](#-run_async_js "-run_async_js的直接链接")

此方法用于以异步方式执行 js 代码。

**参数：**

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `script` | `str` | 必填 | js 脚本文本 |
| `*args` | - | 无 | 传入的参数，按顺序在js文本中对应`arguments[0]`、`arguments[1]`... |
| `as_expr` | `bool` | `False` | 是否作为表达式运行，为`True`时`args`参数无效 |

**返回：**`None`

---

### 📌 `run_cdp()`[​](#-run_cdp "-run_cdp的直接链接")

此方法用于执行 Chrome DevTools Protocol 语句。

cdp 用法详见 [Chrome DevTools Protocol](https://chromedevtools.github.io/devtools-protocol/)。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `cmd` | `str` | 必填 | 协议项目 |
| `**cmd_args` | - | 无 | 项目参数 |

| 返回类型 | 说明 |
| --- | --- |
| `dict` | 执行返回的结果 |

**示例：**

```
# 停止页面加载  
tab.run_cdp('Page.stopLoading')
```

---

### 📌 `run_cdp_loaded()`[​](#-run_cdp_loaded "-run_cdp_loaded的直接链接")

此方法用于执行 Chrome DevTools Protocol 语句，执行前先确保页面加载完毕。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `cmd` | `str` | 必填 | 协议项目 |
| `**cmd_args` | - | 无 | 项目参数 |

| 返回类型 | 说明 |
| --- | --- |
| `dict` | 执行返回的结果 |

---

✅️️ cookies 及缓存[​](#️️-cookies-及缓存 "✅️️ cookies 及缓存的直接链接")
----------------------------------------------------------

### 📌 `set.cookies()`[​](#-setcookies "-setcookies的直接链接")

此方法用于设置 cookie。可设置一个或多个。

设置一个 cookie 支持的格式：

* `Cookie`：单个`Cookie`对象
* `str`：`'name=value; domain=****; ...'`或`'name=****; value=****; domain=****; ...'`格式，只支持用`';'`分隔
* `dict`：`{'name': '****', 'value': '****', 'domain': '****', ...}`或`{name: value, 'domain': '****', ...}`格式

设置多个 cookie 支持的格式：

* `list`或`tuple`：上面几种形式的单个 cookie 放到列表中传入即可
* `dict`：`{name1: value1, name2: value2, ..., 'domain': '****', ...}`格式
* `str`：`'name1=value1; name2=value2; ... domain=****; ...'`格式，多个 cookie 之间只能用`';'`分隔
* `CookieJar`：单个`CookieJar`对象

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `cookies` | `Cookie` `CookieJar` `list` `tuple` `str` `dict` | 必填 | cookies 信息 |

**返回：**`None`

**示例：**

```
# 可以接受多种类型的参数  
cookies1 = ['name1=value1', 'name2=value2']  
cookies2 = 'name1=value1; name2=value2; path=/; domain=.example.com;'  
cookies3 = {'name1': 'value1', 'name2': 'value2', 'domain': '.example.com'}  
tab.set.cookies(cookies1)
```

---

### 📌 `set.cookies.clear()`[​](#-setcookiesclear "-setcookiesclear的直接链接")

此方法用于清除所有 cookie。

**参数：** 无

**返回：**`None`

---

### 📌 `set.cookies.remove()`[​](#-setcookiesremove "-setcookiesremove的直接链接")

此方法用于删除一个 cookie。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `name` | `str` | 必填 | cookie 的 name 字段 |
| `url` | `str` | `None` | cookie 的 url 字段 |
| `domain` | `str` | `None` | cookie 的 domain 字段 |
| `path` | `str` | `None` | cookie 的 path 字段 |

**返回：**`None`

---

### 📌 `set.session_storage()`[​](#-setsession_storage "-setsession_storage的直接链接")

此方法用于设置或删除某项 sessionStorage 信息。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `item` | `str` | 必填 | 要设置的项 |
| `value` | `str` `False` | 必填 | 为`False`时，删除该项 |

**返回：**`None`

**示例：**

```
tab.set.session_storage(item='abc', value='123')
```

---

### 📌 `set.local_storage()`[​](#-setlocal_storage "-setlocal_storage的直接链接")

此方法用于设置或删除某项 localStorage 信息。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `item` | `str` | 必填 | 要设置的项 |
| `value` | `str` `False` | 必填 | 为`False`时，删除该项 |

**返回：**`None`

---

### 📌 `clear_cache()`[​](#-clear_cache "-clear_cache的直接链接")

此方法用于清除缓存，可选择要清除的项。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `session_storage` | `bool` | `True` | 是否清除 sessionstorage |
| `local_storage` | `bool` | `True` | 是否清除 localStorage |
| `cache` | `bool` | `True` | 是否清除 cache |
| `cookies` | `bool` | `True` | 是否清除 cookies |

**返回：**`None`

**示例：**

```
tab.clear_cache(cookies=False)  # 除了 cookies，其它都清除
```

---

  

✅️️ 运行参数设置[​](#️️-运行参数设置 "✅️️ 运行参数设置的直接链接")
-------------------------------------------

各种设置功能藏在`set`属性中。

### 📌 `set.retry_times()`[​](#-setretry_times "-setretry_times的直接链接")

此方法用于设置连接失败时重连次数。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `times` | `int` | 必填 | 次数 |

**返回：**`None`

### 📌 `set.retry_interval()`[​](#-setretry_interval "-setretry_interval的直接链接")

此方法用于设置连接失败时重连间隔。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `interval` | `float` | 必填 | 秒数 |

**返回：**`None`

### 📌 `set.timeouts()`[​](#-settimeouts "-settimeouts的直接链接")

此方法用于设置三种超时时间，单位为秒。可单独设置，为`None`表示不改变原来设置。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `base` | `float` | `None` | 整体超时时间 |
| `page_load` | `float` | `None` | 页面加载超时时间 |
| `script` | `float` | `None` | 脚本运行超时时间 |

**返回：**`None`

**示例：**

```
tab.set.timeouts(base=10, page_load=30)
```

---

### 📌 `set.load_mode`[​](#-setload_mode "-setload_mode的直接链接")

此属性用于设置页面加载策略，调用其方法选择某种策略。

| 方法名称 | 参数 | 说明 |
| --- | --- | --- |
| `normal()` | 无 | 等待页面完全加载完成，为默认状态 |
| `eager()` | 无 | 等待文档加载完成就结束，不等待资源加载 |
| `none()` | 无 | 页面连接完成就结束 |

**示例：**

```
tab.set.load_mode.normal()  
tab.set.load_mode.eager()  
tab.set.load_mode.none()
```

---

### 📌 `set.user_agent()`[​](#-setuser_agent "-setuser_agent的直接链接")

此方法用于为浏览器当前标签页设置 user agent。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `ua` | `str` | 必填 | user agent 字符串 |
| `platform` | `str` | `None` | 平台类型，如`'android'` |

**返回：**`None`

---

### 📌 `set.headers()`[​](#-setheaders "-setheaders的直接链接")

此方法用于设置额外添加到当前页面请求 headers 的参数。

headers 可以是`dict`格式的，也可以是文本格式。

文本格式不同字段用`\n`分隔，字段 key 和 value 用`': '`分隔，即从浏览器直接复制的格式。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `headers` | `dict` `str` | 必填 | headers 信息 |

**返回：**`None`

**示例：**

```
# dict格式  
h = {'connection': 'keep-alive', 'accept-charset': 'GB2312,utf-8;q=0.7,*;q=0.7'}  
tab.set.headers(headers=h)  
  
# 文本格式  
h = '''  
connection: keep-alive  
accept-charset: GB2312,utf-8;q=0.7,*;q=0.7  
'''  
tab.set.headers(headers=h)
```

---

✅️️ 窗口管理[​](#️️-窗口管理 "✅️️ 窗口管理的直接链接")
-------------------------------------

窗口管理功能藏在`set.window`属性中。

### 📌 `set.window.max()`[​](#-setwindowmax "-setwindowmax的直接链接")

此方法用于使窗口最大化。

**参数：** 无

**返回：**`None`

**示例：**

```
tab.set.window.max()
```

---

### 📌 `set.window.mini()`[​](#-setwindowmini "-setwindowmini的直接链接")

此方法用于使窗口最小化。

**参数：** 无

**返回：**`None`

---

### 📌 `set.window.full()`[​](#-setwindowfull "-setwindowfull的直接链接")

此方法用于使窗口切换到全屏模式。

**参数：** 无

**返回：**`None`

---

### 📌 `set.window.normal()`[​](#-setwindownormal "-setwindownormal的直接链接")

此方法用于使窗口切换到普通模式。

**参数：** 无

**返回：**`None`

---

### 📌 `set.window.size()`[​](#-setwindowsize "-setwindowsize的直接链接")

此方法用于设置窗口大小。只传入一个参数时另一个参数不会变化。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `width` | `int` | `None` | 窗口宽度 |
| `height` | `int` | `None` | 窗口高度 |

**返回：**`None`

**示例：**

```
tab.set.window.size(500, 500)
```

---

### 📌 `set.window.location()`[​](#-setwindowlocation "-setwindowlocation的直接链接")

此方法用于设置窗口位置。只传入一个参数时另一个参数不会变化。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `x` | `int` | `None` | 距离顶部距离 |
| `y` | `int` | `None` | 距离左边距离 |

**返回：**`None`

**示例：**

```
tab.set.window.location(500, 500)
```

---

### 📌 `set.window.hide()`[​](#-setwindowhide "-setwindowhide的直接链接")

此方法用于隐藏浏览器窗口。

与 headless 模式不一样，这个方法是直接隐藏浏览器进程。在任务栏上也会消失。只支持 Windows 系统，并且必需已安装 pypiwin32 库才可使用。

不过，窗口隐藏后，如果有新窗口出现，整个浏览器又会显现出来。

**参数：** 无

**返回：**`None`

**示例：**

```
tab.set.window.hide()
```

注意

* 浏览器隐藏后并没有关闭，下次运行程序还会接管已隐藏的浏览器
* 浏览器隐藏后，如果有新建标签页，会自行显示出来

---

### 📌 `set.window.show()`[​](#-setwindowshow "-setwindowshow的直接链接")

此方法用于显示当前浏览器窗口。

**参数：** 无

**返回：**`None`

---

✅️️ 页面滚动[​](#️️-页面滚动 "✅️️ 页面滚动的直接链接")
-------------------------------------

页面滚动的功能藏在`scroll`属性中。

### 📌 `scroll()`或`scroll.down()`[​](#-scroll或scrolldown "-scroll或scrolldown的直接链接")

这两个方法效果是一样的，用于使页面向下滚动若干像素，水平位置不变。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `pixel` | `int` | 必填 | 滚动的像素 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumTab` | `ChromiumTab`执行滚动时返回页面对象自身 |
| `MixTab` | `MixTab`执行滚动时返回页面对象自身 |
| `ChromiumFrame` | `ChromiumFrame`执行滚动时返回页面对象自身 |

---

### 📌 `scroll.up()`[​](#-scrollup "-scrollup的直接链接")

此方法用于使页面向上滚动若干像素，水平位置不变。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `pixel` | `int` | 必填 | 滚动的像素 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumTab` | `ChromiumTab`执行滚动时返回页面对象自身 |
| `MixTab` | `MixTab`执行滚动时返回页面对象自身 |
| `ChromiumFrame` | `ChromiumFrame`执行滚动时返回页面对象自身 |

**示例：**

```
tab.scroll.up(30)
```

---

### 📌 `scroll.right()`[​](#-scrollright "-scrollright的直接链接")

此方法用于使页面向右滚动若干像素，垂直位置不变。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `pixel` | `int` | 必填 | 滚动的像素 |

| 返回类型 |   说明 |
| --- | --- |
| `ChromiumTab` | `ChromiumTab`执行滚动时返回页面对象自身 |
| `MixTab` | `MixTab`执行滚动时返回页面对象自身 |
| `ChromiumFrame` | `ChromiumFrame`执行滚动时返回页面对象自身 |

---

### 📌 `scroll.left()`[​](#-scrollleft "-scrollleft的直接链接")

此方法用于使页面向左滚动若干像素，垂直位置不变。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `pixel` | `int` | 必填 | 滚动的像素 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumTab` | `ChromiumTab`执行滚动时返回页面对象自身 |
| `MixTab` | `MixTab`执行滚动时返回页面对象自身 |
| `ChromiumFrame` | `ChromiumFrame`执行滚动时返回页面对象自身 |

---

### 📌 `scroll.to_top()`[​](#-scrollto_top "-scrollto_top的直接链接")

此方法用于滚动页面到顶部，水平位置不变。

**参数：** 无

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumTab` | `ChromiumTab`执行滚动时返回页面对象自身 |
| `MixTab` | `MixTab`执行滚动时返回页面对象自身 |
| `ChromiumFrame` | `ChromiumFrame`执行滚动时返回页面对象自身 |

**示例：**

```
tab.scroll.to_top()
```

---

### 📌 `scroll.to_bottom()`[​](#-scrollto_bottom "-scrollto_bottom的直接链接")

此方法用于滚动页面到底部，水平位置不变。

**参数：** 无

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumTab` | `ChromiumTab`执行滚动时返回页面对象自身 |
| `MixTab` | `MixTab`执行滚动时返回页面对象自身 |
| `ChromiumFrame` | `ChromiumFrame`执行滚动时返回页面对象自身 |

---

### 📌 `scroll.to_half()`[​](#-scrollto_half "-scrollto_half的直接链接")

此方法用于滚动页面到垂直中间位置，水平位置不变。

**参数：** 无

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumTab` | `ChromiumTab`执行滚动时返回页面对象自身 |
| `MixTab` | `MixTab`执行滚动时返回页面对象自身 |
| `ChromiumFrame` | `ChromiumFrame`执行滚动时返回页面对象自身 |

---

### 📌 `scroll.to_rightmost()`[​](#-scrollto_rightmost "-scrollto_rightmost的直接链接")

此方法用于滚动页面到最右边，垂直位置不变。

**参数：** 无

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumTab` | `ChromiumTab`执行滚动时返回页面对象自身 |
| `MixTab` | `MixTab`执行滚动时返回页面对象自身 |
| `ChromiumFrame` | `ChromiumFrame`执行滚动时返回页面对象自身 |

---

### 📌 `scroll.to_leftmost()`[​](#-scrollto_leftmost "-scrollto_leftmost的直接链接")

此方法用于滚动页面到最左边，垂直位置不变。

**参数：** 无

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumTab` | `ChromiumTab`执行滚动时返回页面对象自身 |
| `MixTab` | `MixTab`执行滚动时返回页面对象自身 |
| `ChromiumFrame` | `ChromiumFrame`执行滚动时返回页面对象自身 |

---

### 📌 `scroll.to_location()`[​](#-scrollto_location "-scrollto_location的直接链接")

此方法用于滚动页面到滚动到指定位置。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `x` | `int` | 必填 | 水平位置，单位是像素 |
| `y` | `int` | 必填 | 垂直位置，单位是像素 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumTab` | `ChromiumTab`执行滚动时返回页面对象自身 |
| `MixTab` | `MixTab`执行滚动时返回页面对象自身 |
| `ChromiumFrame` | `ChromiumFrame`执行滚动时返回页面对象自身 |

**示例：**

```
tab.scroll.to_location(300, 50)
```

---

### 📌 `scroll.to_see()`[​](#-scrollto_see "-scrollto_see的直接链接")

此方法用于滚动页面直到元素可见。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `loc_or_ele` | `str` `tuple` `ChromiumElement` | 必填 | 元素的定位信息，可以是元素、定位符 |
| `center` | `bool` `None` | `None` | 是否尽量滚动到页面正中，为`None`时如果被遮挡，则滚动到页面正中 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumTab` | `ChromiumTab`执行滚动时返回页面对象自身 |
| `MixTab` | `MixTab`执行滚动时返回页面对象自身 |
| `ChromiumFrame` | `ChromiumFrame`执行滚动时返回页面对象自身 |

**示例：**

```
# 滚动到某个已获取到的元素  
ele = tab.ele('tag:div')  
tab.scroll.to_see(ele)  
  
# 滚动到按定位符查找到的元素  
tab.scroll.to_see('tag:div')
```

---

✅️️ 滚动设置[​](#️️-滚动设置 "✅️️ 滚动设置的直接链接")
-------------------------------------

页面滚动有两种方式，一种是滚动时直接跳到目标位置，第二种是平滑滚动，需要一定时间。后者滚动时间难以确定，容易导致程序不稳定，点击不准确的问题。

一些网站会在 css 设置中指定网站使用平滑滚动，这是我们不希望的，但本着让开发者拥有充分选择权利的原则，本库没有强制修改，而是提供两项设置供开发者选择。

### 📌 `set.scroll.smooth()`[​](#-setscrollsmooth "-setscrollsmooth的直接链接")

此方法设置网站是否开启平滑滚动。建议用此方法为网页关闭平滑滚动。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_off` | `bool` | `True` | `bool`表示开或关 |

**返回：**`None`

**示例：**

```
tab.set.scroll.smooth(on_off=False)
```

---

### 📌 `set.scroll.wait_complete()`[​](#-setscrollwait_complete "-setscrollwait_complete的直接链接")

此方法用于设置滚动后是否等待滚动结束。在不想关闭网页平滑滚动功能时，可开启此设置以保障滚动结束后才执行后面的步骤

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_off` | `bool` | `True` | `bool`表示开或关 |

**返回：**`None`

**示例：**

```
tab.set.scroll.wait_complete(on_off=True)
```

---

✅️️ 弹出消息处理[​](#️️-弹出消息处理 "✅️️ 弹出消息处理的直接链接")
-------------------------------------------

### 📌 `handle_alert()`[​](#-handle_alert "-handle_alert的直接链接")

此方法用于处理提示框。  
它能够设置等待时间，等待提示框出现才进行处理，若超时没等到提示框，返回`False`。  
也可只获取提示框文本而不处理提示框。
还可以处理下一个出现的提示框，这在处理离开页面时触发的弹窗非常有用。

注意

程序无法接管一个 已经弹出了提示框的浏览器或标签页。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `accept` | `bool` `None` | `True` | `True`表示确认，`False`表示取消，`None`不会按按钮但依然返回文本值 |
| `send` | `str` | `None` | 处理 prompt 提示框时可输入文本 |
| `timeout` | `float` | `None` | 等待提示框出现的超时时间（秒），为`None`时使用页面整体超时时间 |
| `next_one` | `bool` | `False` | 是否处理下一个出现的弹窗，为`True`时`timeout`参数无效 |

| 返回类型 | 说明 |
| --- | --- |
| `str` | 提示框内容文本 |
| `False` | 未等到提示框则返回`False` |

**示例：**

```
# 确认提示框并获取提示框文本  
txt = tab.handle_alert()  
  
# 点击取消  
tab.handle_alert(accept=False)  
  
# 给 prompt 提示框输入文本并点击确定  
tab.handle_alert(accept=True, send='some text')  
  
# 不处理提示框，只获取提示框文本  
txt = tab.handle_alert(accept=None)
```

---

### 📌 自动处理[​](#-自动处理 "📌 自动处理的直接链接")

标签页对象可使用`set.auto_handle_alert()`方法设置自动处理该 tab 的提示框，使提示框不会弹窗而直接被处理掉。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_off` | `bool` | `True` | 开或关 |
| `accept` | `bool` | `True` | 确定还是取消 |

**返回：**`None`

**示例：**

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.set.auto_handle_alert()  # 这之后出现的弹窗都会自动确认
```

---

### 📌 全局自动处理[​](#-全局自动处理 "📌 全局自动处理的直接链接")

如果需要设置所有标签页都自动处理 alert，可用`Chromium`对象进行设置。

```
from DrissionPage import Chromium  
  
browser = Chromium()  
browser.set.auto_handle_alert()
```

或者

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.browser.set.auto_handle_alert()
```

---

✅️️ 关闭及重连[​](#️️-关闭及重连 "✅️️ 关闭及重连的直接链接")
----------------------------------------

### 📌 `disconnect()`[​](#-disconnect "-disconnect的直接链接")

此方法用于页面对象断开与页面的连接，但不关闭标签页。断开后，对象不能对标签页进行操作。

Tab 和`ChromiumFrame`对象都有此方法。

**参数：** 无

**返回：**`None`

---

### 📌 `reconnect()`[​](#-reconnect "-reconnect的直接链接")

此方法用于关闭与页面连接，然后重建一个新连接。

这主要用于应付长期运行导致内存占用过高，断开连接可释放内存，然后重连继续控制浏览器。

Tab 和`ChromiumFrame`对象都有此方法。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `wait` | `float` | `0` | 关闭后等待多少秒再连接 |

**返回：**`None`

---

### 📌 `close()`[​](#-close "-close的直接链接")

此方法用于关闭标签页。可关闭自己或自己以外的。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `others` | `bool` | `False` | 是否关闭自己以外的标签页 |
| `session` | `bool` | `False` | 是否同时关闭内置`Session`对象，只对自己有效 |

**返回：**`None`

[上一页

🛰️ 访问网页](/browser_control/visit)[下一页

🛰️ 获取网页信息](/browser_control/get_page_info)

* [✅️️ 页面跳转](#️️-页面跳转)
  + [📌 `get()`](#-get)
  + [📌 `back()`](#-back)
  + [📌 `forward()`](#-forward)
  + [📌 `refresh()`](#-refresh)
  + [📌 `stop_loading()`](#-stop_loading)
  + [📌 `set.blocked_urls()`](#-setblocked_urls)
* [✅️️ 元素管理](#️️-元素管理)
  + [📌 `add_ele()`](#-add_ele)
  + [📌 `remove_ele()`](#-remove_ele)
* [✅️️ 执行脚本或命令](#️️-执行脚本或命令)
  + [📌 `run_js()`](#-run_js)
  + [📌 `run_js_loaded()`](#-run_js_loaded)
  + [📌 `run_async_js()`](#-run_async_js)
  + [📌 `run_cdp()`](#-run_cdp)
  + [📌 `run_cdp_loaded()`](#-run_cdp_loaded)
* [✅️️ cookies 及缓存](#️️-cookies-及缓存)
  + [📌 `set.cookies()`](#-setcookies)
  + [📌 `set.cookies.clear()`](#-setcookiesclear)
  + [📌 `set.cookies.remove()`](#-setcookiesremove)
  + [📌 `set.session_storage()`](#-setsession_storage)
  + [📌 `set.local_storage()`](#-setlocal_storage)
  + [📌 `clear_cache()`](#-clear_cache)
* [✅️️ 运行参数设置](#️️-运行参数设置)
  + [📌 `set.retry_times()`](#-setretry_times)
  + [📌 `set.retry_interval()`](#-setretry_interval)
  + [📌 `set.timeouts()`](#-settimeouts)
  + [📌 `set.load_mode`](#-setload_mode)
  + [📌 `set.user_agent()`](#-setuser_agent)
  + [📌 `set.headers()`](#-setheaders)
* [✅️️ 窗口管理](#️️-窗口管理)
  + [📌 `set.window.max()`](#-setwindowmax)
  + [📌 `set.window.mini()`](#-setwindowmini)
  + [📌 `set.window.full()`](#-setwindowfull)
  + [📌 `set.window.normal()`](#-setwindownormal)
  + [📌 `set.window.size()`](#-setwindowsize)
  + [📌 `set.window.location()`](#-setwindowlocation)
  + [📌 `set.window.hide()`](#-setwindowhide)
  + [📌 `set.window.show()`](#-setwindowshow)
* [✅️️ 页面滚动](#️️-页面滚动)
  + [📌 `scroll()`或`scroll.down()`](#-scroll或scrolldown)
  + [📌 `scroll.up()`](#-scrollup)
  + [📌 `scroll.right()`](#-scrollright)
  + [📌 `scroll.left()`](#-scrollleft)
  + [📌 `scroll.to_top()`](#-scrollto_top)
  + [📌 `scroll.to_bottom()`](#-scrollto_bottom)
  + [📌 `scroll.to_half()`](#-scrollto_half)
  + [📌 `scroll.to_rightmost()`](#-scrollto_rightmost)
  + [📌 `scroll.to_leftmost()`](#-scrollto_leftmost)
  + [📌 `scroll.to_location()`](#-scrollto_location)
  + [📌 `scroll.to_see()`](#-scrollto_see)
* [✅️️ 滚动设置](#️️-滚动设置)
  + [📌 `set.scroll.smooth()`](#-setscrollsmooth)
  + [📌 `set.scroll.wait_complete()`](#-setscrollwait_complete)
* [✅️️ 弹出消息处理](#️️-弹出消息处理)
  + [📌 `handle_alert()`](#-handle_alert)
  + [📌 自动处理](#-自动处理)
  + [📌 全局自动处理](#-全局自动处理)
* [✅️️ 关闭及重连](#️️-关闭及重连)
  + [📌 `disconnect()`](#-disconnect)
  + [📌 `reconnect()`](#-reconnect)
  + [📌 `close()`](#-close)

* 🚀 控制浏览器
* 🛰️ Page 对象

本页总览

🛰️ Page 对象
==========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

`ChromiumPage`和`WebPage`是 4.1 之前用于连接和控制浏览器的对象。

4.1 这些功能由`Chromium`实现，但`ChromiumPage`和`WebPage`仍能正常使用。

对比`Chromium`，`ChromiumPage`和`WebPage`在连接浏览器时可以少写一行代码，但在多标签页操作的时候容易造成混乱。

更详细的用法可以看旧版文档。

✅️️ `ChromiumPage`[​](#️️-chromiumpage "️️-chromiumpage的直接链接")
--------------------------------------------------------------

`ChromiumPage`把浏览器管理功能和一个标签页（默认接管时激活那个）控制功能整合在一起。

可看作浏览器对象，但同时控制了一个标签页。

如果项目只需要使用单标签页，用`ChromiumPage`会比较方便。

`ChromiumPage`创建的标签页对象为`ChromiumTab`，没有切换模式功能。

| 初始化参数 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `addr_or_opts` | `str` `int` `ChromiumOptions` | `None` | 浏览器启动配置或接管信息。 传入 'ip: port' 字符串、端口数字或`ChromiumOptions`对象时按配置启动或接管浏览器； 为`None`时使用配置文件配置启动浏览器 |
| `tab_id` | `str` | `None` | 要控制的标签页 id，不指定默认为激活的 |

```
from DrissionPage import ChromiumPage  
  
page = ChromiumPage()  
page.get('http://DrissionPage.cn')  
print(page.title)
```

---

  

✅️️ `WebPage`[​](#️️-webpage "️️-webpage的直接链接")
-----------------------------------------------

`WebPage`覆盖了`ChromiumPage`所有功能，并且增加了切换模式功能，创建的标签页对象为`MixTab`。

| 初始化参数 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `mode` | `str` | `'d'` | 运行模式，可选`'d'`或`'s'` |
| `chromium_options` | `bool` `ChromiumOptions` | `None` | `ChromiumOptions`对象，传入`None`时从默认 ini 文件读取，传入`False`时不读取 ini 文件，使用默认配置 |
| `session_or_options` | `SessionOptions` `None` `False` | `None` | `Session`对象或`SessionOptions`对象，传入`None`时从默认 ini 文件读取，传入`False`时不读取 ini 文件，使用默认配置 |

```
from DrissionPage import WebPage  
  
page = WebPage()  
page.get('http://DrissionPage.cn')  
print(page.title)  
page.change_mode()  
print(page.title)
```

[上一页

🛰️ 上传文件](/browser_control/upload)[下一页

🛩️ 概述](/SessionPage/intro)

* [✅️️ `ChromiumPage`](#️️-chromiumpage)
* [✅️️ `WebPage`](#️️-webpage)

* 🚀 控制浏览器
* 🛰️ 截图和录像

本页总览

🛰️ 截图和录像
========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

✅️️ 页面截图[​](#️️-页面截图 "✅️️ 页面截图的直接链接")
-------------------------------------

使用页面对象的`get_screenshot()`方法对页面进行截图，可对整个网页、可见网页、指定范围截图。

对可视范围外截图需要 90 以上版本浏览器支持。

下面三个参数三选一，优先级：`as_bytes`>`as_base64`>`path`。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `path` | `str` `Path` | `None` | 保存图片的路径，为`None`时保存在当前文件夹 |
| `name` | `str` | `None` | 完整文件名，后缀可选`'jpg'`、`'jpeg'`、`'png'`、`'webp'`，为`None`时以用 jpg 格式 |
| `as_bytes` | `str` `True` | `None` | 是否以字节形式返回图片，可选`'jpg'`、`'jpeg'`、`'png'`、`'webp'`、`None`、`True` 不为`None`时`path`参数无效 为`True`时选用 jpg 格式 |
| `as_base64` | `str` `True` | `None` | 是否以 base64 形式返回图片，可选`'jpg'`、`'jpeg'`、`'png'`、`'webp'`、`None`、`True` 不为`None`时`path`参数无效 为`True`时选用 jpg 格式 |
| `full_page` | `bool` | `False` | 是否整页截图，为`True`截取整个网页，为`False`截取可视窗口 |
| `left_top` | `Tuple[int, int]` | `None` | 截取范围左上角坐标 |
| `right_bottom` | `Tuple[int, int]` | `None` | 截取范围右下角坐标 |

| 返回类型 | 说明 |
| --- | --- |
| `bytes` | `as_bytes`生效时返回图片字节 |
| `str` | `as_bytes`和`as_base64`为`None`时返回图片完整路径 |
| `str` | `as_base64`生效时返回 base64 格式的字符串 |

说明

如`path`为包含文件名的完整路径，`name`参数无效。

**示例：**

```
# 对整页截图并保存  
tab.get_screenshot(path='tmp', name='pic.jpg', full_page=True)
```

️️ ✅️️ 元素截图[​](#️️-️️-元素截图 "️️ ✅️️ 元素截图的直接链接")
----------------------------------------------

使用元素对象的`get_screenshot()`方法对元素进行截图。

若元素范围超出视口，需 90 以上版本内核支持。

下面三个参数三选一，优先级：`as_bytes`>`as_base64`>`path`。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `path` | `str` `Path` | `None` | 保存图片的路径，为`None`时保存在当前文件夹 |
| `name` | `str` | `None` | 完整文件名，后缀可选`'jpg'`、`'jpeg'`、`'png'`、`'webp'`，为`None`时以用 jpg 格式 |
| `as_bytes` | `str` `True` | `None` | 是否以字节形式返回图片，可选`'jpg'`、`'jpeg'`、`'png'`、`'webp'`、`None`、`True` 不为`None`时`path`和`as_base64`参数无效 为`True`时选用 jpg 格式 |
| `as_base64` | `str` `True` | `None` | 是否以 base64 形式返回图片，可选`'jpg'`、`'jpeg'`、`'png'`、`'webp'`、`None`、`True` 不为`None`时`path`参数无效 为`True`时选用 jpg 格式 |
| `scroll_to_center` | `bool` | `True` | 截图前是否滚动到视口中央 |

| 返回类型 | 说明 |
| --- | --- |
| `bytes` | `as_bytes`生效时返回图片字节 |
| `str` | `as_bytes`和`as_base64`为`None`时返回图片完整路径 |
| `str` | `as_base64`生效时返回 base64 格式的字符串 |

说明

如`path`为包含文件名的完整路径，`name`参数无效。

**示例：**

```
img = tab('tag:img')  
img.get_screenshot()  
bytes_str = img.get_screenshot(as_bytes='png')  # 返回截图二进制文本
```

---

✅️️ 页面录像[​](#️️-页面录像 "✅️️ 页面录像的直接链接")
-------------------------------------

使用页面对象的`screencast`功能，可以录取屏幕图片或视频。

### 📌 设置录制模式[​](#-设置录制模式 "📌 设置录制模式的直接链接")

录制模式一共有 5 种，通过`screencast.set_mode.xxx_mode()`设置。

| 模式 | 说明 |
| --- | --- |
| `video_mode()` | 持续录制页面，停止时生成没有声音的视频 |
| `frugal_video_mode()` | 页面有变化时才录制，停止时生成没有声音的视频 |
| `js_video_mode()` | 可生成有声音的视频，但需要手动启动 |
| `imgs_mode()` | 持续对页面进行截图 |
| `frugal_imgs_mode()` | 页面有变化时才保存页面图像 |

### 📌 设置存放路径[​](#-设置存放路径 "📌 设置存放路径的直接链接")

使用`screencast.set_save_path()`设置录制结果保存路径。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `save_path` | `str` `Path` | `None` | 保存图片或视频的路径 |

**返回：**`None`

### 📌 `screencast.start()`[​](#-screencaststart "-screencaststart的直接链接")

此方法用于开始录制浏览器窗口。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `save_path` | `str` `Path` | `None` | 保存图片或视频的路径 |

**返回：**`None`

注意

保存路径必需设置，无论是用`screencast.set()`还是`screencast.start()`都可以。

### 📌 `screencast.stop()`[​](#-screencaststop "-screencaststop的直接链接")

此方法用于停止录取屏幕。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `video_name` | `str` | `None` | 视频文件名，为`None`时以当前时间名命 |
| `suffix` | `str` | `'mp4'` | 视频文件后缀 |
| `coding` | `str` | `'mp4v'` | 视频编码格式，仅`video_mode`模式有效，根据`cv2.VideoWriter_fourcc()`定义 |

| 返回类型 | 说明 |
| --- | --- |
| `str` | 保存为视频时返回视频文件路径，否则返回保存图片的文件夹路径 |

### 📌 注意事项[​](#-注意事项 "📌 注意事项的直接链接")

* 使用`video_mode`和`frugal_video_mode`时，保存路径和保存文件名必需是英文。
* 使用`video_mode`和`frugal_video_mode`时，需先安装 opencv 库。`pip install opencv-python`
* 使用`js_video_mode`时，需用鼠标手动选择要录制的目标，才能开始录制
* 使用`js_video_mode`时，如要对一个窗口进行录制，需在另一个窗口开始录制，否则如窗口出现跳转，会使录制失效

### 📌 示例[​](#-示例 "📌 示例的直接链接")

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.screencast.set_save_path('video')  # 设置视频存放路径  
tab.screencast.set_mode.video_mode()  # 设置录制  
tab.screencast.start()  # 开始录制  
tab.wait(3)  
tab.screencast.stop()  # 停止录制
```

[上一页

🛰️ 获取控制台信息](/browser_control/console)[  下一页

🛰️ 上传文件](/browser_control/upload)

* [✅️️ 页面截图](#️️-页面截图)
* [️️ ✅️️ 元素截图](#️️-️️-元素截图)
* [✅️️ 页面录像](#️️-页面录像)
  + [📌 设置录制模式](#-设置录制模式)
  + [📌 设置存放路径](#-设置存放路径)
  + [📌 `screencast.start()`](#-screencaststart)
  + [📌 `screencast.stop()`](#-screencaststop)
  + [📌 注意事项](#-注意事项)
  + [📌 示例](#-示例)

* 🚀 控制浏览器
* 🛰️ 标签页管理

本页总览

🛰️ 标签页管理
========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

浏览器的标签页由 Tab 对象（`ChromiumTab`和`MixTab`）控制。

与网页的交互都由标签页对象进行。

默认情况下，一个标签页由一个 Tab 对象控制。

多个 Tab 对象可以同时操作，不需要切换焦点，也不需要激活到前台。

提醒

当禁用单例模式后，一个标签页也可以被多个 Tab 对象同时控制。

✅️️ 获取标签页对象[​](#️️-获取标签页对象 "✅️️ 获取标签页对象的直接链接")
----------------------------------------------

### 📌 获取最后激活的标签页[​](#-获取最后激活的标签页 "📌 获取最后激活的标签页的直接 链接")

`Chromium`对象的`latest_tab`属性返回最后激活的标签页对象。

说明

如`Settings.singleton_tab_obj`为`True`，此属性返回标签页对象的 tab id。

```
from DrissionPage import Chromium  
  
browser = Chromium()  
tab = browser.latest_tab  # 获取最新标签页对象
```

---

### 📌 获取指定标签页[​](#-获取指定标签页 "📌 获取指定标签页的直接链接")

`Chromium`对象的`get_tab()`和`get_tabs()`方法用于获取指定的标签页对象。

可指定标签页序号、id、标题、url、类型等条件用于检索。api 详见 “浏览器对象” 章节。

说明

* 当`id_or_num`不为`None`时，其它参数无效
* `title`、`url`和`tab_type`三个参数是与关系
* 如传入序号，序号与标签页视觉排序不一定一致，而是按照激活顺序排列。

```
from DrissionPage import Chromium  
  
browser = Chromium()  
tab1 = browser.get_tab(1)  # 获取列表中第一个标签页的对象  
tab2 = browser.get_tab('5399F4ADFE3A27503FFAA56390344EE5')  # 获取指定id的标签页对象  
tab3 = browser.get_tab(url='DrissionPage.cn')  # 获取第一个url中带 'DrissionPage.cn' 的标签页对象  
tabs = browser.get_tabs(url='DrissionPage.cn')  # 获取所有url中带 'DrissionPage.cn' 的标签页对象
```

注意

Tab 对象默认为单例，即一个实体标签页只有一个`MixTab`对象。`get_tab()`返回的标签页可能是同一个。

---

### 📌 新建标签页并获取对象[​](#-新建标签页 并获取对象 "📌 新建标签页并获取对象的直接链接")

`Chromium`对象的`new_tab()`方法用于新建一个标签页，返回其对象。

```
from DrissionPage import Chromium  
  
browser = Chromium()  
browser.new_tab(url='http://DrissionPage.cn')
```

说明

当传入`url`参数时，程序会根据`load_mode`设置访问页面，除了`none`模式，都将等待页面加载完毕。
如果新建多个标签页不想等待，可批量新建不传入`url`参数的标签页，再遍历使用`get()`。

---

### 📌 获取点击后出现的标签页[​](#-获取点击后出现的标签页 "📌 获取点击后出现的标签页的直接链接")

在预期点击元素会出现新标签页时，可用元素的`click.for_new_tab()`方法实行点击，点击后会返回新标 签页对象。

具体参数见元素交互章节。

可直接运行以下示例：

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.get('https://DrissionPage.cn')  
ele = tab.ele('.wwads-cn wwads-horizontal').ele('tag:img')  
if ele:  
    tab2 = ele.click.for_new_tab()  # 点击并获取新tab对象  
    tab2.set.activate()  
    ele2 = tab2.ele('确认访问', timeout=5)  
    if ele2:  
        ele2.wait(.5).click()  
else:  
    print('支持开源作者，请关闭广告屏蔽功能，谢谢。')
```

元素对象的`click.middle()`方法可用中键点击`<a>`元素，可强制在新标签页打开链接，并返回新标签页对象。

可直接运行以下示例：

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.get('https://DrissionPage.cn')  
ele = tab.ele('.wwads-cn wwads-horizontal').ele('tag:img')  
if ele:  
    tab2 = ele.click.middle()  # 中键点击元素，并获取新tab对象  
    tab2.set.activate()  
    ele2 = tab2.ele('确认访问', timeout=5)  
    if ele2:  
        ele2.wait(.5).click()  
else:  
    print('支持开源作者，请关闭广告屏蔽功能，谢谢。')
```

---

✅️️ 多标签页协同[​](#️️-多标签页协同 "✅️️ 多标签页协同的直接链接")
-------------------------------------------

这个示例在一个标签页中遍历列表元素，点击打开新标签页，获取信息后关闭。

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.get('https://gitee.com/explore/all')  
  
links = tab.eles('t:h3')  
for link in links[:-1]:  
    # 点击链接并获取新标签页对象  
    new_tab = link.click.for_new_tab()  
    # 等待新标签页加载  
    new_tab.wait.load_start()  
    # 打印标签页标题  
    print(new_tab.title)  
    # 关闭新打开的标签页  
    new_tab.close()
```

✅️️ 使用多例[​](#️️-使用多例 "✅️️ 使用多例的直接链接")
-------------------------------------

默认情况下，Tab 对象是单例的，即一个标签页只有一个对象，即使重复使用`get_tab()`，获取的都是同一个对象。

这主要是防止新手不理解机制，反复创建多个连接导致资源耗费。

实际上允许多个 Tab 对象同时操作一个标签页，每个负责不同的工  作。比如一个执行主逻辑流程，另外的监视页面，处理各种弹窗。

要允许多例，可用`Settings`设置：

```
from DrissionPage.common import Settings  
  
Settings.set_singleton_tab_obj(False)
```

**示例**

```
from DrissionPage import Chromium  
from DrissionPage.common import Settings  
  
browser = Chromium()  
browser.new_tab()  
browser.new_tab()  
  
# 未启用多例：  
tab1 = browser.get_tab(1)  
tab2 = browser.get_tab(1)  
print(id(tab1), id(tab2))  
  
# 启用多例：  
Settings.set_singleton_tab_obj(False)  
tab1 = browser.get_tab(1)  
tab2 = browser.get_tab(1)  
print(id(tab1), id(tab2))
```

**输出：**

```
2347582903056 2347582903056  
2347588741840 2347588877712
```

可见第一次输出两个 Tab 对象是同一个，第二次输出是独立的。

[上一页

🛰️ 浏览器对象](/browser_control/browser_object)[下一页

🛰️ 访问网页](/browser_control/visit)

* [✅️️ 获取标签页对象](#️️-获取标签页对象)
  + [📌 获取最后激活的标签页](#-获取最后激活的标签页)
  + [📌 获取指定标签页](#-获取指定标签页)
  + [📌 新建标签页并获取对象](#-新建标签页并获取对象)
  + [📌 获取点击后出现的标签页](#-获取点击后出现的标签页)
* [✅️️ 多标签页协同](#️️-多标签页协同)
* [✅️️ 使用多例](#️️-使用多例)

* 🚀 控制浏览器
* 🛰️ 上传文件

本页总览

🛰️ 上传文件
=======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

上传文件有两种方式：

* 拦截文件输入框，自动填入路径
* 找到`<input>`元素，填入文件路径

✅️️ 自然的交互[​](#️️-自然的交互 "✅️️ 自然的交互的直接链接")
----------------------------------------

传统自动化工具的文件上传，需要开发者在 DOM 里找到文件上传控件，然后用元素对象的`input()`方法填入路径。

有些上传控件是临时加载的，有些藏得很深，找起来费时费力。

本库提供一种自然的文件上传方式，无需在 DOM 里找控件，只要自然地点击触发文件选择框，程序就能主动截获，并填写设定好的路径，开发更省事。

### 📌 `click.to_upload()`[​](#-clickto_upload "-clickto_upload的直接链接")

浏览器元素对象拥有此方法，用于上传文件到网页。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `file_paths` | `str` `Path` `list` `tuple` | 必填 | 文件路径，如果上传框支持多文件，可传入列表或字符串，字符串时多个文件用`\n`分隔 |
| `by_js` | `bool` | `False` | 是否用 js 方式点击，逻辑与`click()`一致 |

**返回：**`None`

**示例**

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
ele = tab('#uploadButton')  
ele.click.to_upload(r'C:\text.txt')
```

---

### 📌 手动方式[​](#-手动方式 "📌 手动方式的直接链接")

上面的方法使用默认点击方式触发上传，假如页面要求其它触发方式，可自行手动写上传逻辑。

**步骤：**

* 设置要上传的文件路径，多路径传入`list`、`tuple`或以`\n`分隔的字符串
* 点击会触发文件选择框的按钮
* 调用等待录入语句，确保输入完整

**示例：**

```
# 设置要上传的文件路径  
tab.set.upload_files('demo.txt')  
# 点击触发文件选择框按钮  
btn_ele.click()  
# 等待路径填入  
tab.wait.upload_paths_inputted()
```

点击按钮后，文本选择框被拦截不会弹出，但可以看到文件路径已经传入其中。

由于此动作是异步输入，需显式等待输入完成才进行下一步操作。

---

### 📌 注意事项[​](#-注意事项 "📌 注意事项的直接链接")

如果您要操作的上传控件在一个异域的`<iframe>`，那必需用这个`<iframe>`对象来设置上传路径，而不能用页面对象设置。

❌ 错误做法：

```
tab.set.upload_paths('demo.txt')  
tab.get_frame(1).ele('@type=file').click()  
tab.wait.upload_paths_inputted()
```

⭕ 正确做法：

```
iframe = tab.get_frame(1)  
iframe.set.upload_paths('demo.txt')  
iframe.ele('@type=file').click()  
iframe.wait.upload_paths_inputted()
```

如果`<iframe>`和主页面是同域的，则用域名对象和`<iframe>`对象设置均可。

---

✅️️ 传统方式[​](#️️-传统方式 "✅️️ 传统方式的直接链接")
-------------------------------------

传统方  式，需要开发者在 DOM 里找到文件上传控件，然后用元素对象的`input()`方法填入路径。

文件上传控件是`type`属性为`'file'`的`<input>`元素进行输入，把文件路径输入到元素即可，用法与输入文本一致。

稍有不同的是，无论`clear`参数是什么，都会清空原控件内容。

如果控件支持多文件上传，多个路径用`list`、`tuple`或以`\n`分隔的字符串传入。

```
upload = tab('tag:input@type=file')  
  
# 传入一个路径  
upload.input('D:\\test1.txt')  
  
# 传入多个路径，方式 1  
paths = 'D:\\test1.txt\nD:\\test2.txt'  
upload.input(paths)  
  
# 传入多个路径，方式 2  
paths = ['D:\\test1.txt', 'D:\\test2.txt']  
upload.input(paths)
```

如果`<input>`元素很好找，这种方式是很简便的。

有些`<input>`是临时加载的，或者经过修饰隐藏很深，找起来很费劲。

万一有些上传是用 js 控制的，这种方式未必能奏效。

[上一页

🛰️ 截图和录像](/browser_control/screen)[下一页

🛰️ Page 对象](/browser_control/pages)

* [✅️️ 自然的交互](#️️-自然的交互)
  + [📌 `click.to_upload()`](#-clickto_upload)
  + [📌 手动方式](#-手动方式)
  + [📌 注意事项](#-注意事项)
* [✅️️ 传统方式](#️️-传统方式)

* 🚀 控制浏览器
* 🛰️ 访问网页

本页总览

🛰️ 访问网页
=======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

本节介绍 Tab 对象访问网页的相关内容。

✅️️ 连接方法[​](#️️-连接方法 "✅️️ 连接方法的直接链接")
-------------------------------------

### 📌 `get()`[​](#-get "-get的直接链接")

该方法用于跳转到一个网址。当连接失败时，程序会进行重试。

可指定本地文件路径。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `url` | `str` | 必填 | 目标 url，可指向本地文件路径 |
| `show_errmsg` | `bool` | `False` | 连接出错时是否显示和抛出异常 |
| `retry` | `int` | `None` | 重试次数，为`None`时使用页面参数，默认`3` |
| `interval` | `float` | `None` | 重试间隔（秒），为`None`时使用页面参数，默认`2` |
| `timeout` | `float` | `None` | 加载超时时间（秒） |
| ------- | ------- | --- | ------ 以下参数仅 s 模式有效 ------ |
| `params` | `dict` | `None` | url 请求参数 |
| `data` | `dict` `str` | `None` | 携带的数据 |
| `json` | `dict` `str` | `None` | 要发送的 JSON 数据，会自动设置 Content-Type 为`'application/json'` |
| `headers` | `dict` | `None` | 请求头 |
| `cookies` | `dict` `CookieJar` | `None` | cookies 信息 |
| `files` | `Any` | `None` | 要上传的文件，可以是一个字典，其中键是文件名，值是文件对象或文件路径 |
| `auth` | `Any` | `None` | 身份认证信息 |
| `allow_redirects` | `bool` | `True` | 是否允许重定向 |
| `proxies` | `dict` | `None` | 代理信息 |
| `hooks` | `Any` | `None` | 回调方法 |
| `stream` | `bool` | `None` | 是否使用流式传输 |
| `verify` | `bool` `str` | `None` | 是否验证 SSL 证书 |
| `cert` | `str` `Tuple[str, str]` | `None` | SSL 客户端证书文件的路径(.pem 格式)，或('cert', 'key')元组 |

| 返回类型 | 说明 |
| --- | --- |
| `bool` | 访问是否成功 |

**示例：**

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.get('http://DrissionPage.cn')
```

---

### 📌 `post()`[​](#-post "-post的直接链接")

此方法用内置的`Session`对象以 POST 方式发送请求。

因为`post()`是使用`requests`的`post()`方法发送请求，参数和用法与`requests`一致。

此方法返回请求结果`Response`对象。

s 模式时，`post()`后结果还可用页面对象的`html`或`json`属性获取。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `url` | `str` | 必填 | 目标 url，可指向本地文件路径 |
| `show_errmsg` | `bool` | `False` | 连接出错时是否显示和抛出异常 |
| `retry` | `int` | `None` | 重试次数，为`None`时使用页面参数，默认`3` |
| `interval` | `float` | `None` | 重试间隔（秒），为`None`时使用页面参数，默认`2` |
| `timeout` | `float` | `None` | 加载超时时间（秒） |
| `params` | `dict` | `None` | url 请求参数 |
| `data` | `dict` `str` | `None` | 携带的数据 |
| `json` | `dict` `str` | `None` | 要发送的 JSON 数据，会自动设置 Content-Type 为`'application/json'` |
| `headers` | `dict` | `None` | 请求头 |
| `cookies` | `dict` `CookieJar` | `None` | cookies 信息 |
| `files` | `Any` | `None` | 要上传的文件，可以是一个字典，其中键是文件名，值是文件对象或文件路径 |
| `auth` | `Any` | `None` | 身份认证信息 |
| `allow_redirects` | `bool` | `True` | 是否允许重定向 |
| `proxies` | `dict` | `None` | 代理信息 |
| `hooks` | `Any` | `None` | 回调方法 |
| `stream` | `bool` | `None` | 是否使用流式传输 |
| `verify` | `bool` `str` | `None` | 是否验证 SSL 证书 |
| `cert` | `str` `Tuple[str, str]` | `None` | SSL 客户端证书文件的路径(.pem 格式)，或('cert', 'key')元组 |

| 返回类型 | 说明 |
| --- | --- |
| `Response` | 获取到的`Response`对象 |

---

✅️️ 设置超时和重试[​](#️️-设置超时和重试 "✅️️ 设置超时和重试的直接链接")
----------------------------------------------

网络不稳定时，访问页面不一定成功，`get()`方法内置了超时和重试功能。通过`retry`、`interval`、`timeout`三个参数进行设置。  
其中，如不指定`timeout`参数，该参数会使用`ChromiumPage`的`timeouts`属性的`page_load`参数中的值。

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.get('http://DrissionPage.cn', retry=1, interval=1, timeout=1.5)
```

---

✅️️ 加载模式[​](#️️-加载模式 "✅️️ 加载模式的直接链接")
-------------------------------------

### 📌 概述[​](#-概述 "📌 概述的直接链接")

加载模式是指 d 模式下程序在页面加载阶段的行为模式，有以下三种：

* `normal()`：常规模式，会等待页面加载完毕，超时自动重试或停止，默认使用此模式
* `eager()`：加载完 DOM 或超时即停止加载，不加载页面资源
* `none()`：超时也不会自动停止，除非加载完成

前两种模式下，页面加载过程会阻塞程序，直到加载完毕才执行后面的操作。

`none()`模式下，只在连接阶段阻塞程序，加载阶段可自行根据情况执行`stop_loading()`停止加载。

这样提供给用户非常大的自由度，可等到关键数据包或元素出现就主动停止页面加载，大幅提升执行效率。

注意

加载完成是指主文档完成，并不包括由 js 触发的加载和重定向的加载。
当文档加载完成，程序就判断加载完毕，此后发生的重定向或 js 加载数据需用其它逻辑处理。

**示例：**

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.set.load_mode.eager()  # 设置为eager模式  
tab.get('http://DrissionPage.cn')
```

---

### 📌 模式设置[​](#-模式设置 "📌 模式设置的直接链接")

可通过 ini 文件、`ChromiumOptions`对象和页面对象的`set.load_mode.****()`方法进行设置。

运行时可随时动态设置。

**配置对象中设置**

```
from DrissionPage import ChromiumOptions, Chromium  
  
co = ChromiumOptions().set_load_mode('none')  
browser = Chromium(co)
```

**运行中设置**

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.set.load_mode.none()
```

---

### 📌 `none`模式技巧[​](#-none模式技巧 "-none模式技巧的直接链接")

**示例 1，配合监听器**

跟监听器配合，可在获取到需要的数据包时，主动停止加载。

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.set.load_mode.none()  # 设置加载模式为none  
  
tab.listen.start('api/getkeydata')  # 指定监听目标并启动监听  
tab.get('http://www.hao123.com/')  # 访问网站  
packet = tab.listen.wait()  # 等待数据包  
tab.stop_loading()  # 主动停止加载  
print(packet.response.body)  # 打印数据包正文
```

**示例 2，配合元素查找**

跟元素查找配合，可在获取到某个指定元素时，主动停止加载。

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.set.load_mode.none()  # 设置加载模式为none  
  
tab.get('http://www.hao123.com/')  # 访问网站  
ele = tab.ele('中国日报')  # 查找text包含“中国日报”  的元素  
tab.stop_loading()  # 主动停止加载  
print(ele.text)  # 打印元素text
```

**示例 2，配合页面特征**

可等待到页面到达某种状态时，主动停止加载。比如多级跳转的登录，可等待 title 变化到最终目标网址时停止。

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.set.load_mode.none()  # 设置加载模式为none  
  
tab.get('http://www.hao123.com/')  # 访问网站  
tab.wait.title_change('hao123')  # 等待title变化出现目标文本  
tab.stop_loading()  # 主动停止加载
```

[上一页

🛰️ 标签页管理](/browser_control/tabs)[下一页

🛰️ 页面交互](/browser_control/page_operation)

* [✅️️ 连接方法](#️️-连接方法)
  + [📌 `get()`](#-get)
  + [📌 `post()`](#-post)
* [✅️️ 设置超时和重试](#️️-设置超时和重试)
* [✅️️ 加载模式](#️️-加载模式)
  + [📌 概述](#-概述)
  + [📌 模式设置](#-模式设置)
  + [📌 `none`模式技巧](#-none模式技巧)

* 🚀 控制浏览器
* 🛰️ 等待

本页总览

🛰️ 等待
=====

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

网络环境不稳定，页面 js 运行时间也难以确定，自动化过程中经常遇到需要等待的情况。

如果总是用`sleep()`，显得不太优雅，等待多了浪费时间，等待不够会导致报错。

因此，程序能够智能等待是非常重要的，DrissionPage 内置了一些等待方法，可以提高程序稳定性和效率。

它们藏在页面对象和元素对象的`wait`属性里。

等待方法均有`timeout`参数，可自行设得超时时间，也可以设置超时后返回`False`还是抛出异常。

✅️️ 浏览器对象的等待方法[​](#️️-浏览器对象的等待方法 "✅️️ 浏览器对象的等待方法的直接链接")
-------------------------------------------------------

### 📌 `wait.new_tab()`[​](#-waitnew_tab "-waitnew_tab的直接链接")

此方法用于等待新标签页出现。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` | `None` | 超时时间（秒），为`None`时使用页面`timeout`设置 |
| `curr_tab` | `str` `ChromiumTab` `MixTab` | `None` | 指定当前最新的 Tab 对象或标签页 id，用于判断新标签页出现，为`None`自动获取 |
| `raise_err` | `bool` | `None` | 等待失败时是否报错，为`None`时根据`Settings`设置 |

| 返回类型 | 说明 |
| --- | --- |
| `str` | 等待成返回新标签页 id |
| `False` | 等待失败返回`False` |

---

### 📌 `wait.download_begin()`[​](#-waitdownload_begin "-waitdownload_begin的直接链接")

此方法用于等待浏览器一个下载任务开始，详见下载功能章节。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` | `None` | 超时时间（秒），为`None`时使用页面`timeout`设置 |
| `cancel_it` | `bool` | `False` | 是否取消该任务 |

| 返回类型 | 说明 |
| --- | --- |
| `DownloadMission` | 等待成功返回下载任务对象 |
| `False` | 等待失败 |

**示例：**

```
tab('#download_btn').click()  # 点击按钮触发下载  
tab.wait.download_begin()  # 等待下载开始
```

---

### 📌 `wait.downloads_done()`[​](#-waitdownloads_done "-waitdownloads_done的直接链接")

此方法用于等待浏览器所有下载任务完成，详见下载功能章节。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` | `None` | 超时时间（秒），为`None`时无限等待 |
| `cancel_if_timeout` | `bool` | `True` | 超时时是否取消剩余任务 |

| 返回类型 | 说明 |
| --- | --- |
| `bool` | 是否等待成功 |

---

✅️️ 页面对象的等待方法[​](#️️-页面对象的等待方法 "✅️️ 页面对象的等待方法的直接链接")
----------------------------------------------------

页面对象指`ChromiumTab`、`MixTab`和`ChromiumFrame`。

**用法：**

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.get('http://DrissionPage.cn')  
tab.wait.ele_displayed('tag:div')
```

### 📌 `wait.load_start()`[​](#-waitload_start "-waitload_start的直接链接")

此方法用于等待页面进入加载状态。  
我们经常会通过点击元素进入下一个网页，并立刻获取新页面的元素。  
但若跳转前的页面拥有和跳转后页面相同定位符的元素，会导致过早获取元素，跳转后失效的问题。  
使用此方法，会阻塞程序，等待页面开始加载后再继续，从而避免上述问题。  
我们通常只需等待页面加载开始，程序会自动等待加载结束。

注意

`get()`已内置等待加载开始，后无须跟`wait.load_start()`。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` `None` `True` | `None` | 超时时间（秒），为`None`或`Ture`时使用页面`timeout`设置 为数字时等待相应时间 |
| `raise_err` | `bool` | `None` | 等待失败时是否报错，为`None`时根据`Settings`设置 |

| 返回类型 | 说明 |
| --- | --- |
| `bool` | 等待结束时是否进入加载状态 |

**示例：**

```
ele.click()  # 点击某个元素  
tab.wait.load_start()  # 等待页面进入加载状态  
# 执行在新页面的操作  
print(page.title)
```

---

### 📌 `wait.doc_loaded()`[​](#-waitdoc_loaded "-waitdoc_loaded的直接链接")

此方法用于等待页面文档加载完成。  
一般来说都无需开发者使用，程序大部分动作都会自动等待加载完成再执行。

注意

* 此功能仅用于等待页面主 document 加载，不能用于等待 js 加载的变化。
* 除非`load_mode`为`None`，`get()`方法已内置等待加载完成，后面无须添加等待。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` `None` `True` | `None` | 超时时间（秒），为`None`或`Ture`时使用页面`timeout`设置 为数字时等待相应时间 |
| `raise_err` | `bool` | `None` | 等待失败时是否报错，为`None`时根据`Settings`设置 |

| 返回类型 | 说明 |
| --- | --- |
| `bool` | 等待结束时是否完成加载完成 |

---

### 📌 `wait.eles_loaded()`[​](#-waiteles_loaded "-waiteles_loaded的直接链接")

此方法用于等待元素被加载到 DOM，可等待全部或任意一个加载。  
有时一个元素的正常出现是下一步操作的前提，用此方法可以防止一些元素加载速度慢于程序动作速度导致的误操作。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` `list` | 必填 | 要等待的元素，定位符 |
| `timeout` | `float` | `None` | 超时时间（秒），为`None`时使用页面`timeout`设置 |
| `any_one` | `bool` | `False` | 是否等待到一个就返回 |
| `raise_err` | `bool` | `None` | 等待失败时是否报错，为`None`时根据`Settings`设置 |

| 返回类型 | 说明 |
| --- | --- |
| `bool` | 是否等待成功 |

**示例：**

```
ele1.click()  # 点击某个元素  
page.wait.eles_loaded('#div1')  # 等待 id 为 div1 的元素加载  
ele2.click()  # div1 加载完成后再执行下一步操作
```

---

### 📌 `wait.ele_displayed()`[​](#-waitele_displayed "-waitele_displayed的直接链接")

此方法用于等待一个元素变成显示状态。  
如果当前 DOM 中查找不到指定元素，则会自动等待元素加载，再等待它显示。  
元素隐藏是指元素在 DOM 内，但处于隐藏状态（即使在视口内且不被遮挡）。  
父元素隐藏时子元素也是隐藏的。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `loc_or_ele` | `str` `Tuple[str, str]` `ChromiumElement` | 必填 | 要等待的元素，可以是元素或定位符 |
| `timeout` | `float` | `None` | 超时时间（秒），为`None`时使用页面`timeout`设置 |
| `raise_err` | `bool` | `None` | 等待失败时是否报错，为`None`时根据`Settings`设置 |

| 返回类型 | 说明 |
| --- | --- |
| `bool` | 是否等待成功 |

**示例：**

```
# 等待 id 为 div1 的元素显示，超时使用页面设置  
tab.wait.ele_displayed('#div1')  
  
# 等待 id 为 div1 的元素显示，设置超时3秒  
tab.wait.ele_displayed('#div1', timeout=3)  
  
# 等待已获取到的元素被显示  
ele = tab.ele('#div1')  
tab.wait.ele_displayed(ele)
```

---

### 📌 `wait.ele_hidden()`[​](#-waitele_hidden "-waitele_hidden的直接链接")

此方法用于等待一个元素变成隐藏状态。  
如果当前 DOM 中查找不到指定元素，则会自动等待元素加载，再等待它隐藏。  
元素隐藏是指元素在 DOM 内，但处于隐藏状态（即使在视口内且不被遮挡）。  
父元素隐藏时子元素也是隐藏的。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `loc_or_ele` | `str` `Tuple[str, str]` `ChromiumElement` | 必填 | 要等待的元素，可以是元素或定位符 |
| `timeout` | `float` | `None` | 超时时间（秒），为`None`时使用页面`timeout`设置 |
| `raise_err` | `bool` | `None` | 等待失败时是否报错，为`None`时根据`Settings`设置 |

| 返回类型 | 说明 |
| --- | --- |
| `bool` | 是否等待成功 |

---

### 📌 `wait.ele_deleted()`[​](#-waitele_deleted "-waitele_deleted的直接链接")

此方法用于等待一个元素被从 DOM 中删除。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `loc_or_ele` | `str` `Tuple[str, str]` `ChromiumElement` | 必填 | 要等待的元素，可以是元素或定位符 |
| `timeout` | `float` | `None` | 超时时间（秒），为`None`时使用页面`timeout`设置 |
| `raise_err` | `bool` | `None` | 等待失败时是否报错，为`None`时根据`Settings`设置 |

| 返回类型 | 说明 |
| --- | --- |
| `bool` | 是否等待成功 |

---

### 📌 `wait.download_begin()`[​](#-waitdownload_begin-1 "-waitdownload_begin-1的直接链接")

此方法用于等待下载开始，详见下载功能章节。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` | `None` | 超时时间（秒），为`None`时使用页面`timeout`设置 |
| `cancel_it` | `bool` | `False` | 是否取消该任务 |

| 返回类型 | 说明 |
| --- | --- |
| `DownloadMission` | 等待成功返回下载任务对象 |
| `False` | 等待失败 |

**示例：**

```
tab('#download_btn').click()  # 点击按钮触发下载  
tab.wait.download_begin()  # 等待下载开始
```

---

### 📌 `wait.downloads_done()`[​](#-waitdownloads_done-1 "-waitdownloads_done-1的直接链接")

此方法用于等待本标签页所有下载任务完成，详见下载功能章节。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` | `None` | 超时时间（秒），为`None`时无限等待 |
| `cancel_if_timeout` | `bool` | `True` | 超时时是否取消剩余任务 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumTab` | `ChromiumTab`对象等待成功返回自己 |
| `MixTab` | `MixTab`对象等待成功返回自己 |
| `False` | 等待失败 |

---

### 📌 `wait.upload_paths_inputted()`[​](#-waitupload_paths_inputted "-waitupload_paths_inputted的直接链接")

此方法用于等待自动填写上传文件路径。详见文件上传章节。

**参数：** 无

**返回：**`None`

**示例：**

```
# 设置要上传的文件路径  
tab.set.upload_files('demo.txt')  
# 点击触发文件选择框按钮  
btn_ele.click()  
# 等待路径填入  
tab.wait.upload_paths_inputted()
```

---

### 📌 `wait.title_change()`[​](#-waittitle_change "-waittitle_change的直接链接")

此方法用于等待 title 变成包含或不包含指定文本。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `text` | `str` | 必填 | 用于识别的文本 |
| `exclude` | `bool` | `False` | 是否排除，为`True`时当 title 不包含`text`指定文本时返回`True` |
| `timeout` | `bool` | `float` | 超时时间（秒） |
| `raise_err` | `bool` | `None` | 等待失败时是否报错，为`None`时根据`Settings`设置 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumTab` | 等待成功`ChromiumTab`对象返回自身 |
| `MixTab` | 等待成功`MixTab`对象返回自身 |
| `ChromiumFrame` | `<iframe>`元素的等待返回对象自身 |
| `False` | 等待失败 |

---

### 📌 `wait.url_change()`[​](#-waiturl_change "-waiturl_change的直接链接")

此方法用于等待 url 变成包含或不包含指定文本。  
比如有些网站登录时会进行多重跳转，url 发生多次变化，可用此功能等待到达最终需要的页面。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `text` | `str` | 必填 | 用于识别的文本 |
| `exclude` | `bool` | `False` | 是否排除，为`True`时当 url 不包含`text`指定文本时返回`True` |
| `timeout` | `bool` | `float` | 超时时间（秒） |
| `raise_err` | `bool` | `None` | 等待失败时是否报错，为`None`时根据`Settings`设置 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumTab` | 等待成功`ChromiumTab`对象返回自身 |
| `MixTab` | 等待成功`MixTab`对象返回自身 |
| `ChromiumFrame` | `<iframe>`元素的等待返回对象自身 |
| `False` | 等待失败 |

**示例：**

```
# 访问网站  
tab.get('https://www.*****.cn/login/')  # 访问登录页面  
tab.ele('#username').input('***')  # 执行登录逻辑  
tab.ele('#password').input('***\n')  
  
tab.wait.url_change('https://www.*****.cn/center/')  # 等待url变成后台url
```

---

### 📌 `wait.alert_closed()`[​](#-waitalert_closed "-waitalert_closed的直接链接")

此方法用于等待弹出框被关闭。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `bool` | `float` | 超时时间（秒），为`None`无限等待 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumTab` | 等待成功`ChromiumTab`对象返回自身 |
| `MixTab` | 等待成功`MixTab`对象返回自身 |
| `False` | 等待失败 |

---



---

  

✅️️ 元素对象的等待方法[​](#️️-元素对象的等待方法 "✅️️ 元素对象的等待方法的直接链接")
----------------------------------------------------

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.get('http://DrissionPage.cn')  
ele = tab('tag:div')  
ele.wait.covered()
```

### 📌 `wait.displayed()`[​](#-waitdisplayed "-waitdisplayed的直接链接")

此方法用于等待元素从隐藏状态变成显示状态。  
元素隐藏是指元素在 DOM 内，但处于隐藏状态（即使在视口内且不被遮挡）。父元素隐藏时子元素也是隐藏的。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` | `None` | 等待超时时间（秒），为`None`则使用元素所在页面超时时间 |
| `raise_err` | `bool` | `None` | 等待失败时是否报错，为`None`时根据`Settings`设置 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 元素对象自身 |
| `ChromiumFrame` | `<iframe>`元素的等待返回对象自身 |
| `False` | 等待失败 |

**示例：**

```
# 等待元素显示，超时使用ele所在页面设置  
ele.wait.displayed()
```

---

### 📌 `wait.hidden()`[​](#-waithidden "-waithidden的直接链接")

此方法用于等待元素从显示状态变成隐藏状态。  
元素隐藏是指元素在 DOM 内，但处于隐藏状态（即使在视口内且不被遮挡）。父元素隐藏时子元素也是隐藏的。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` | `None` | 等待超时时间（秒），为`None`则使用元素所在页面超时时间 |
| `raise_err` | `bool` | `None` | 等待失败时是否报错，为`None`时根据`Settings`设置 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 元素对象自身 |
| `ChromiumFrame` | `<iframe>`元素的等待返回对象自身 |
| `False` | 等待失败 |

**示例：**

```
# 等待元素不显示，超时为3秒  
ele.wait.hidden(timeout=3)
```

---

### 📌 `wait.deleted()`[​](#-waitdeleted "-waitdeleted的直接链接")

此方法用于等待元素被从 DOM 删除。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` | `None` | 等待超时时间（秒），为`None`则使用元素所在页面超时时间 |
| `raise_err` | `bool` | `None` | 等待失败时是否报错，为`None`时根据`Settings`设置 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 元素对象自身 |
| `ChromiumFrame` | `<iframe>`元素的等待返回对象自身 |
| `False` | 等待失败 |

**示例：**

```
# 等待元素显示，超时使用ele所在页面设置  
ele.wait.deleted()
```

---

### 📌 `wait.has_rect()`[​](#-waithas_rect "-waithas_rect的直接链接")

此方法用于等待元素被赋予大小。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` | `None` | 等待超时时间（秒），为`None`则使用元素所在页面超时时间 |
| `raise_err` | `bool` | `None` | 等待失败时是否报错，为`None`时根据`Settings`设置 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 元素对象自身 |
| `ChromiumFrame` | `<iframe>`元素的等待返回对象自身 |
| `False` | 等待失败 |

---

### 📌 `wait.covered()`[​](#-waitcovered "-waitcovered的直接链接")

此方法用于等待元素被其它元素覆盖。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` | `None` | 等待超时时间（秒），为`None`则使用元素所在页面超时时间 |
| `raise_err` | `bool` | `None` | 等待失败时是否报错，为`None`时根据`Settings`设置 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 元素对象自身 |
| `ChromiumFrame` | `<iframe>`元素的等待返回对象自身 |
| `False` | 等待失败 |

---

### 📌 `wait.not_covered()`[​](#-waitnot_covered "-waitnot_covered的直接链接")

此方法用于等待元素不被其它元素覆盖。  
可用于等待遮挡被操作元素的“加载中”遮罩消失。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` | `None` | 等待超时时间（秒），为`None`则使用元素所在页面超时时间 |
| `raise_err` | `bool` | `None` | 等待失败时是否报错，为`None`时根据`Settings`设置 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 元素对象自身 |
| `ChromiumFrame` | `<iframe>`元素的等待返回对象自身 |
| `False` | 等待失败 |

---

### 📌 `wait.enabled()`[​](#-waitenabled "-waitenabled的直接链接")

此方法用于等待元素变为可用状态。  
不可用状态的元素仍然在 DOM 内，`disabled`属性为`False`。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` | `None` | 等待超时时间（秒），为`None`则使用元素所在页面超时时间 |
| `raise_err` | `bool` | `None` | 等待失败时是否报错，为`None`时根据`Settings`设置 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 元素对象自身 |
| `ChromiumFrame` | `<iframe>`元素的等待返回对象自身 |
| `False` | 等待失败 |

---

### 📌 `wait.disabled()`[​](#-waitdisabled "-waitdisabled的直接链接")

此方法用于等待元素变为不可用状态。  
不可用状态的元素仍然在 DOM 内，`disabled`属性为`True`。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` | `None` | 等待超时时间（秒），为`None`则使用元素所在页面超时时间 |
| `raise_err` | `bool` | `None` | 等待失败时是否报错，为`None`时根据`Settings`设置 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 元素对象自身 |
| `ChromiumFrame` | `<iframe>`元素的等待返回对象自身 |
| `False` | 等待失败 |

---

### 📌 `wait.stop_moving()`[​](#-waitstop_moving "-waitstop_moving的直接链接")

此方法用于等待元素运动结束。如果元素没有大小和位置信息，会在超时时抛出`NoRectError`异常。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` | `None` | 等待超时时间（秒），为`None`则使用元素所在页面超时时间 |
| `gap` | `float` | `0.1` | 检测运动的间隔时间 |
| `raise_err` | `bool` | `None` | 等待失败时是否报错，为`None`时根据`Settings`设置 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 元素对象自身 |
| `ChromiumFrame` | `<iframe>`元素的等待返回对象自身 |
| `False` | 等待失败 |

```
# 等待元素稳定  
tab.ele('#button1').wait.stop_moving()  
# 点击元素  
tab.ele('#button1').click()
```

---

### 📌 `wait.clickable()`[​](#-waitclickable "-waitclickable的直接链接")

此方法用于等待元素可被点击。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `wait_moved` | `bool` | `True` | 是否等待元素运动结束 |
| `timeout` | `float` | `None` | 等待超时时间（秒），为`None`则使用元素所在页面超时时间 |
| `raise_err` | `bool` | `None` | 等待失败时是否报错，为`None`时根据`Settings`设置 |

| 返回类型 | 说明 |
| --- | --- |
| `ChromiumElement` | 元素对象自身 |
| `ChromiumFrame` | `<iframe>`元素的等待返回对象自身 |
| `False` | 等待失败 |

---

### 📌 `wait.disabled_or_deleted()`[​](#-waitdisabled_or_deleted "-waitdisabled_or_deleted的直接链接")

此方法用于等待元素变为不可用或被删除。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` | `None` | 等待超时时间（秒），为`None`则使用元素所在页面超时时间 |
| `raise_err` | `bool` | `None` | 等待失败时是否报错，为`None`时根据`Settings`设置 |

| 返回类型 | 说明 |
| --- | --- |
| `bool` | 是否等待成功 |

---

✅️️ 共有的等待方法[​](#️️-共有的等待方法 "✅️️ 共有的等待方法的直接链接")
----------------------------------------------

### 📌 `wait()`[​](#-wait "-wait的直接链接")

此方法用于等待若干秒。所有对象的等待都可使用这个方法。  
`scope`为`None`时，效果与`time.sleep()`没有区别，等待指定秒数。  
`scope`不为`None`时，获取两个参数之间的一个随机值，等待这个数值的秒数。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `second` | `float` | 必填 | 要等待的秒数，`scope`不为`None`时表示随机数范围起始值 |
| `scope` | `float` | `None` | 随机数范围结束值 |

**返回：** 调用者自身

**示例：**

```
from DrissionPage import Chromium  
  
browser = Chromium()  
  
# 强制等待1秒  
browser.wait(1)  
  
# 获取3.5至8.5之间的一个随机数，等待这个数值的秒数  
browser.wait(3.5, 8.5)
```

[上一页

🛰️ 模式切换](/browser_control/mode_change)[下一页

🛰️ 监听网络数据](/browser_control/listener)

* [✅️️ 浏览器对象的等待方法](#️️-浏览器对象的等待方法)
  + [📌 `wait.new_tab()`](#-waitnew_tab)
  + [📌 `wait.download_begin()`](#-waitdownload_begin)
  + [📌 `wait.downloads_done()`](#-waitdownloads_done)
* [✅️️ 页面对象的等待方法](#️️-页面对象的等待方法)
  + [📌 `wait.load_start()`](#-waitload_start)
  + [📌 `wait.doc_loaded()`](#-waitdoc_loaded)
  + [📌 `wait.eles_loaded()`](#-waiteles_loaded)
  + [📌 `wait.ele_displayed()`](#-waitele_displayed)
  + [📌 `wait.ele_hidden()`](#-waitele_hidden)
  + [📌 `wait.ele_deleted()`](#-waitele_deleted)
  + [📌 `wait.download_begin()`](#-waitdownload_begin-1)
  + [📌 `wait.downloads_done()`](#-waitdownloads_done-1)
  + [📌 `wait.upload_paths_inputted()`](#-waitupload_paths_inputted)
  + [📌 `wait.title_change()`](#-waittitle_change)
  + [📌 `wait.url_change()`](#-waiturl_change)
  + [📌 `wait.alert_closed()`](#-waitalert_closed)
* [✅️️ 元素对象的等待方法](#️️-元素对象的等待方法)
  + [📌 `wait.displayed()`](#-waitdisplayed)
  + [📌 `wait.hidden()`](#-waithidden)
  + [📌 `wait.deleted()`](#-waitdeleted)
  + [📌 `wait.has_rect()`](#-waithas_rect)
  + [📌 `wait.covered()`](#-waitcovered)
  + [📌 `wait.not_covered()`](#-waitnot_covered)
  + [📌 `wait.enabled()`](#-waitenabled)
  + [📌 `wait.disabled()`](#-waitdisabled)
  + [📌 `wait.stop_moving()`](#-waitstop_moving)
  + [📌 `wait.clickable()`](#-waitclickable)
  + [📌 `wait.disabled_or_deleted()`](#-waitdisabled_or_deleted)
* [✅️️ 共有的等待方法](#️️-共有的等待方法)
  + [📌 `wait()`](#-wait)

## <a name="下载"></a>下载

* 🧰 进阶使用
* ⬇️ 下载文件
* ⤵️ download方法

本页总览

⤵️ download方法
=============

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

DrissionPage 每种页面对象都内置一个下载工具，提供任务管理、多线程并发、大文件分块、自动重连、文件名冲突处理等功能。

该工具现已独立打包成一个库，名为 DownloadKit，详细介绍见：[DownloadKit](http://drissionpage.cn/DownloadKitDocs)。

这里只介绍其主要功能，具体使用和设置方法请移步该文档。

✅️️ 功能简介[​](#️️-功能简介 "✅️️ 功能简介的直接链接")
-------------------------------------

### 📌 支持该工具的对象[​](#-支持该工具的对象 "📌 支持该工具的对象的直接链接")

以下对象均支持

* `SessionPage`
* `ChromiumTab`
* `MixTab`
* `ChromiumFrame`
* `ChromiumPage`
* `WebPage`

---

### 📌 下载器功能[​](#-下载器功能 "📌 下载器功能的直接链接")

* 可下载指定 url 文件
* 支持多线程并发下载多个文件
* 大文件自动分块使用多线程下载
* 可对现有文件追加数据
* 自动创建目标路径
* 下载时支持文件重命名
* 自动处理文件名冲突
* 自动去除路径和文件名中非法字符
* 支持 post 方式
* 支持自定义连接参数
* 任务失败自动重试

注意

`DownloadKit`是对 requests 封装实现的，不是调用浏览器功能。
如果下载目标对 headers、data 等有要求，必需手动添加。

---

✅️️ 添加任务[​](#️️-添加任务 "✅️️ 添加任务的直接链接")
-------------------------------------

### 📌 单线程任务[​](#-单线程任务 "📌 单线程任务的直接链接")

使用`download()`方法可添加单线程任务，该方法  是阻塞式的，且只使用一个线程。

**示例：**

```
from DrissionPage import SessionPage  
  
page = SessionPage()  
url = 'https://www.baidu.com/img/flexible/logo/pc/result.png'  
save_path = r'C:\download'  
  
res = page.download(url, save_path)  
print(res)
```

显示：

```
url：https://www.baidu.com/img/flexible/logo/pc/result.png  
文件名：result.png  
目标路径：C:\download  
100% 下载完成 C:\download\result.png  
  
('success', 'C:\\download\\result.png')
```

---

### 📌 并发任务[​](#-并发任务 "📌 并发任务的直接链接")

使用`download.add()`添加并发任务。

**示例：**

```
url1 = 'https://dldir1.qq.com/qqfile/qq/TIM3.4.8/TIM3.4.8.22092.exe'  
url2 = 'https://dldir1.qq.com/qqfile/qq/PCQQ9.7.16/QQ9.7.16.29187.exe'  
save_path = 'files'  
  
page = SessionPage()  
page.download.add(url1, save_path)  
page.download.add(url2, save_path)
```

---

### 📌 文件分块并行下载[​](#-文件分块并行下载 "📌 文件分块并行下载的直接链接")

使用`download.add()`方法的`split`参数可设置大文件是否分块下载。

使用`download.set.block_size()`方法可设置分块大小。

默认情况下载，超过 50M 的文件会自动分块下载。

**示例：**

```
page = SessionPage()  
page.download.set.block_size('30m')  # 设置分块大小  
page.download.add('http://****/demo.zip')  # 默认分块下载  
page.download.add('http://****/demo.zip', split=False)  # 不使用分块下载
```

---

### 📌 阻塞式多线程任务[​](#-阻塞式多线程任务 "📌 阻塞式多线程任务的直接链接")

使用并行分块下载时，也可以使任务逐个下载，在`add()`后使用`wait()`即可。

**示例：**

```
page = SessionPage()  
page.download.add('http://****/demo.zip').wait()  
page.download.add('http://****/demo.zip').wait()
```

---

### 📌 详细使用文档[​](#-详细使用文档 "📌 详细使用文档的直接链接")

以上仅是普通示例，详细功能请查阅：[DownloadKit 添加任务](http://drissionpage.cn/DownloadKitDocs/usage/add_missions/)

---

  

✅️️ 下载设置[​](#️️-下载设置 "✅️️ 下载设置的直接链接")
-------------------------------------

### 📌 全局设置[​](#-全局设置 "📌 全局设置的直接链接")

使用`download.set.****()`方法，可对默认下载行为进行设置。

包括以下设置：

* 保存路径
* 允许使用的线程总数
* 是否启用分块下载
* 分块大小
* 连接失败重试次数
* 重试间隔
* 连接超时时间
* 文件名冲突时的处理方式
* 日志和显示相关设置

---

### 📌 每个任务单独设置[​](#-每个任务单独设置 "📌 每个任务单独设置的直接链接")

新建任务时，`download()`和`add()`方法的参数可对当前任务进行参数设置，覆盖全局设置。

详见上文添加参数的文档。

---

### 📌 详细使用文档[​](#-详细使用文档-1 "📌 详细使用文档的直接链接")

详细设置功能请查阅：[DownloadKit 运行设置](http://drissionpage.cn/DownloadKitDocs/usage/settings/)

---

✅️️ 任务管理[​](#️️-任务管理 "✅️️ 任务管理的直接链接")
-------------------------------------

### 📌 任务对象[​](#-任务对象 "📌 任务对象的直接链接")

对象`Mission`用于管理任务，有以下功能：

* 查看任务状态、信息、进度
* 保存任务参数，如 url、连接参数等
* 取消进行中的任务
* 删除已下载的文件

---

### 📌 获取单个任务对象[​](#-获取单个任务对象 "📌 获取单个任务对象的直接链接")

使用`download.add()`添加任务时，会返回一 个任务对象。

**示例：**

```
mission = page.download.add('http://****.pdf')  
print(mission.id)  # 获取任务id  
print(mission.rate)  # 打印下载进度（百分比）  
print(mission.state)  # 打印任务状态  
print(mission.info)  # 打印任务信息  
print(mission.result)  # 打印任务结果
```

除添加任务时获取对象，也可以使用`download.get_mission()`获取。在上一个示例中可以看到，任务对象有`id`属性，把任务的`id`传入此方法，会返回该任务对象。

**示例：**

```
mission_id = mission.id  
mission = page.download.get_mission(mission_id)
```

---

### 📌 获取全部任务对象[​](#-获取全部任务对象 "📌 获取全部任务对象的直接链接")

使用页面对象的`download.missions`属性，可以获取所有下载任务。该属性返回一个`dict`，保存了所有下载任务。以任务对象的`id`为 key。

```
page.download_set.save_path(r'D:\download')  
page.download('http://****/****1.pdf')  
page.download('http://****/****1.pdf')  
print(page.download.missions)
```

**输出：**

```
{  
    1: <Mission 1 D:\download\xxx1.pdf xxx1.pdf>  
    2: <Mission 2 D:\download\xxx1_1.pdf xxx1_1.pdf>  
    ...  
}
```

---

### 📌 获取下载失败的任务[​](#-获取下载失败的任务 "📌 获取下载失败的任务的直接链接")

使用`download.get_failed_missions()`方法，可以获取下载失败的任务列表。

```
page.download_set.save_path(r'D:\download')  
page.download('http://****/****1.pdf')  
page.download('http://****/****1.pdf')  
print(page.download.get_failed_missions()
```

**输出：**

```
[  
    <Mission 1 状态码：404 None>,  
    <Mission 2 状态码：404 None>  
    ...  
]
```

Tips

获取失败任务对象后，可从其`data`属性读取任务内容，以便记录日志或择机重试。

---

### 📌 详细使用文档[​](#-详细使用文档-2 "📌 详细使用文档的直接链接")

详细设置功能请查阅：[DownloadKit 任务管理](http://drissionpage.cn/DownloadKitDocs/usage/misssions/)

[上一页

⤵️ 概述](/download/intro)[下一页

⤵️ 浏览器下载](/download/browser)

* [✅️️ 功能简介](#️️-功能简介)
  + [📌 支持该工具的对象](#-支持该工具的对象)
  + [📌 下载器功能](#-下载器功能)
* [✅️️ 添加任务](#️️-添加任务)
  + [📌 单线程任务](#-单线程任务)
  + [📌 并发任务](#-并发任务)
  + [📌 文件分块并行下载](#-文件分块并行下载)
  + [📌 阻塞式多线程任务](#-阻塞式多线程任务)
  + [📌 详细使用文档](#-详细使用文档)
* [✅️️ 下载设置](#️️-下载设置)
  + [📌 全  局设置](#-全局设置)
  + [📌 每个任务单独设置](#-每个任务单独设置)
  + [📌 详细使用文档](#-详细使用文档-1)
* [✅️️ 任务管理](#️️-任务管理)
  + [📌 任务对象](#-任务对象)
  + [📌 获取单个任务对象](#-获取单个任务对象)
  + [📌 获取全部任务对象](#-获取全部任务对象)
  + [📌 获取下载失败的任务](#-获取下载失败的任务)
  + [📌 详细使用文档](#-详细使用文档-2)

* 🧰 进阶使用
* ⬇️ 下载文件
* ⤵️ 浏览器下载

本页总览

⤵️ 浏览器下载
========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

本节介绍对浏览器下载任务进行设置的功能。

✅️ 概述[​](#️-概述 "✅️ 概述的直接链接")
----------------------------

### 📌 功能[​](#-功能 "📌 功能的直接链接")

DrissionPage 提供以下功能，用于对浏览器下载任务进行控制：

* 每个 tab 对象可独立设置文件保存路径
* 下载前可指定文件名称，实现文件重命名
* 可设置存在同名文件时的处理方式
* 可获取任务下载进度
* 可等待下载任务结束
* 可取消任务
* 可拦截下载任务并获取其信息

---

⚠️ 注意事项[​](#️-注意事项 "⚠️ 注意事项的直接链接")
----------------------------------

### 📌 记得等待任务结束[​](#-记得等待任务结束 "📌 记得等待任务结束的直接链接")

因技术原因，程序在下载结束时才能对其重命名，在这之前文件名是临时的任务 id。

因此必需等待下载完毕，文件名才能正确命名。无论是否指定文件名都一样。

**示例：**

```
tab = Chromium().latest_tab  
tab('#button').click()  # 点击下载按钮  
tab.wait.download_begin()  # 等待下载开始  
tab.wait.downloads_done()  # 等待所有任务结束
```

---

### 📌 多 Tab 操作时推荐设置临时路径[​](#-多-tab-操作时推荐设置临时路径 "📌 多 Tab 操作时推荐设置临时路径的直接链接")

程序需要把任务下载到一个指定位置，完成后再移动到目标路径。

因此，如果程序涉及多个 Tab 触发下载任务，最好给`Chromium`对象设置一个下载路径。

即使每个 Tab 对象都设置了自己的路径。

**示例：**

```
from DrissionPage import Chromium  
  
browser = Chromium()  
browser.set.download_path('tmp')  # 设置总路径  
  
tab1 = browser.get_tab(1)  
tab1.set.download_path('path1')  
  
tab2 = browser.get_tab(2)  
tab2.set.download_path('path2')
```

---

### 📌 启动下载管理功能[​](#-启动下载管理功能 "📌 启动下载管理功能的直接链接")

本节介绍的下载管理功能默认不开启，此时触发下载任务和手动操作没有区别。

当启动配置中设置了下载路径，或调用`set.download_path()`方法时，管理功能才会启动。

使用`click.to_download()`方法会自动自动此功能。

---

✅️ `click.to_download()`[​](#️-clickto_download "️-clickto_download的直接链接")
--------------------------------------------------------------------------

当预期点击元素后会触发下载，可使用此方法返回下载任务。

使用时可以同时设置下载路径、指定文件名名称。

注意

* 有些下载任务是点击后弹出新标签页，在新标签页触发下载，此时必须设置`new_tab=True`
* 点击后等待下载触发时间为页面对象`timeout`属性（默认10秒），如果需要更久的等待，用下文介绍的方法。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `save_path` | `str` `Path` | 必填 | 保存路径，为`None`保存在原来设置的，如未设置保存到当前路径 |
| `rename` | `str` | `None` | 重命名文件名，为`None`则不修改 |
| `suffix` | `str` | `None` | 指定文件后缀，为`None`则不修改 |
| `new_tab` | `bool` | `False` | 预期的下载是否在新标签页中触发 |
| `by_js` | `bool` | `False` | 是否用 js 方式点击，逻辑与`click()`一致 |
| `timeout` | `float` | `None` | 超时时间（秒），为`None`时使用页面对象默认超时时间 |

| 返回类型 | 说明 |
| --- | --- |
| `DownloadMission` | 下载任务对象 |

**示例：**

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab   
tab.get('https://im.qq.com/pcqq/index.shtml')  
ele = tab('全新体验版下载')  
ele.wait.has_rect()  
mission = ele.click.to_download(save_path='tmp', rename='QQ.exe')  
mission.wait()
```

`click.to_download()`方法能够应付多数情况，不是点击触发的下载任务或更复杂的情况，请根据下文介绍的配置方式使用。

---

✅️ 设置下载路径[​](#️-设置下载路径 "✅️ 设置下载路径的直接链接")
----------------------------------------

### 📌 设置总下载路径[​](#-设置总下载路径 "📌 设置总下载路径的直接链接")

使用`Chromium`对象的`set.download_path()`方法设置下载路径。不设置时，默认下载到程序当前路径。

`Chromium`对象设置下载路径后，后续新建的 Tab 对象均会使用该路径，之前建立的 Tab 对象使用的路径则不会改变。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `path` | `str` `Path` | 必填 | 下载路径 |

**返回：**`None`

**示例：**

```
from DrissionPage import Chromium  
  
browser = Chromium()  
browser.set.download_path(r'C:\tmp')
```

---

### 📌 设置 Tab 下载路径[​](#-设置-tab-下载路径 "📌 设置 Tab 下载路径的直接链接")

使用方法与设置`Chromium`的一致，但只在当前 Tab 对象生效。

**示例：**

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.set.download_path(r'C:\tmp1')  # 设置Tab下载路径
```

---

  

✅️ 设置文件名[​](#️-设置文件名 "✅️ 设置文件名的直接链接")
-------------------------------------

使用`download_file_name()`方法，可在下载前设置文件名，实现下载文件的重命名。

设置的文件名可以不带后缀，程序会根据下载的文件自动补充后缀。

如设置的文件名带`'.'`，且后缀与网络文件不一致，程序会以网络文件的后缀为准。

如果想修改后缀名，设置`suffix`参数即可。

每次触发下载后，该设置会被清空。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `name` | `str` | `None` | 文件名 |
| `suffix` | `str` | `None` | 文件后缀名，传入`''`可去除后缀 |

**返回：**`None`

**示例：**

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.set.download_file_name('new_file')  
tab('t:a').click()  # 点击一个会触发下载的链接  
mission = tab.wait.download_begin()  
mission.wait()  # 记得等待任务触发和结束
```

---

✅️ 等待[​](#️-等待 "✅️ 等待的直接链接")
----------------------------

### 📌 等待下载开始[​](#-等待下载开始 "📌 等待下载开始的直接链接")

点击下载链接后，下载并不会瞬间触发，需要进行等待，才能将其捕获。

使用`wait.download_begin()`方法等待下载开始。

一般来说，标签页触发的下载任务用 Tab 对象进行等待，未被控制的标签页触发的下载，可由`Chromium`对象进行等待。

`cancel_it`参数为`True`时，捕获到任务时会将其取消，以便将返回的下载信息用于其它需要。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` | `None` | 超时时间（秒），`None`为页面默认等待时间 |
| `cancel_it` | `bool` | `False` | 是否取消该任务 |

| 返回类型 | 说明 |
| --- | --- |
| `DownloadMission` | 等待成功且`cancel_it`为`False`时返回下载任务对象 |
| `dict` | 等待成功且`cancel_it`为`True`时返回下载任务信息 |
| `False` | 等待失败返回`False` |

**示例：**

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab('t:a').click()  # 点击一个会触发下载的链接  
tab.wait.download_begin()
```

---

### 📌 等待所有下载任务结束[​](#-等待所有下载任务结束 "📌 等待所有下载任务结束的直接链接")

用`Chromium`对象的`wait.downloads_done()`方法可等待浏览器所有下载任务结束。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` | `None` | 超时时间（秒），为`None`时无限等待 |
| `cancel_if_timeout` | `bool` | `False` | 如超时是否取消未完成的任务 |

| 返回类型 | 说明 |
| --- | --- |
| `bool` | 是否等待成功 |

---

### 📌 等待某 Tab 所有下载任务结束[​](#-等待某-tab-所有下载任务结束 "📌 等待某 Tab 所有下载任务结束的直接链接")

用 Tab 对象的`wait.downloads_done()`方法可等待该 Tab 对象触发点下载任务结束。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `timeout` | `float` | `None` | 超时时间（秒），为`None`时无限等待 |
| `cancel_if_timeout` | `bool` | `False` | 如超时是否取消未完成的任务 |

| 返回类型 | 说明 |
| --- | --- |
| `bool` | 是否等待成功 |

---

✅️ 拦截下载任务[​](#️-拦截下载任务 "✅️ 拦截下载任务的直接链接")
----------------------------------------

`wait.download_begin()`方法有个`cancel_it`参数，当为`True`时，会取消下载任务。

此 时可使用该方法返回的任务信息进行下一步操作，如改用`download()`方法下载等。

**示例：**

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab('t:a').click()  
data = tab.wait.download_begin(cancel_it=True)  
tab.download(data.url)
```

---

✅️ 同名文件的处理[​](#️-同名文件的处理 "✅️ 同名文件的处理的直接链接")
-------------------------------------------

下载遇到  同名文件时，可选择三种处理方式：自动重命名、覆盖、跳过。

使用`set.when_download_file_exists('****')`进行设置。

其中`****`可选`'rename'`、`'overwrite'`、`'skip'`。

也可选择它们的首字母`'r'`、`'o'`、`'s'`。

### 📌 自动重命名[​](#-自动重命名 "📌 自动重命名的直接链接")

设置方法：`set.when_download_file_exists('rename')`

这种方式遇到已有同名文件时会自动将新文件重命名，方式是在后面加上序号。

假设保存路径已存在名为 'abc.zip' 的文件，再下载一个 'abc.zip' 时，新文件会自动重命名为 'abc\_1.zip'。

之后再下载会命名为 'abc\_2.zip'，如此类推。

---

### 📌 覆盖已有文件[​](#-覆盖已有文件 "📌 覆盖已有文件的直接链接")

设置方法：`set.when_download_file_exists('overwrite')`

这种方式会把原有的同名文件替换成新下载的。

---

### 📌 跳过[​](#-跳过 "📌 跳过的直接链接")

设置方法：`set.when_download_file_exists('skip')`

这种方式下，发现已有同名文件时会取消下载任务。

---

✅️ 任务管理[​](#️-任务管理 "✅️ 任务管理的直接链接")
----------------------------------

`wait.download_begin()`方法会返回`DownloadMission`对象，用于浏览器下载任务的管理。

### 📌 获取任务信息[​](#-获取任务信息 "📌 获取任务信息的直接链接")

可获取任务状态、进度、保存路径、文件名等信息。

| 属性名称 | 类型 | 说明 |
| --- | --- | --- |
| `url` | `str` | 返回任务网址 |
| `tab_id` | `str` | 触发任务的 Tab 对象 id |
| `id` | `str` | 任务 id |
| `folder` | `str` | 保存文件夹路径 |
| `name` | `str` | 文件名 |
| `tmp_path` | `str` | 临时文件保存路径 |
| `state` | `str` | 任务状态，'running', 'done', 'canceled', 'skipped' |
| `total_bytes` | `int` | 总字节数 |
| `received_bytes` | `int` | 已接收字节数 |
| `final_path` | `str` `None` | 最终完整路径，任务完成后才产生 |

**示例：**

实时打印任务进度。

```
mission = tab.wait.download_begin()  
while not mission.is_done:  
    print(f'\r{mission.rate}%', end='')
```

---

### 📌 等待任务结束[​](#-等待任务结束 "📌 等待任务结束的直接链接")

使用`DownloadMission`对象的`wait()`方法，可等待任务结束。

| 参数名称 | 类型 | 默认  值 | 说明 |
| --- | --- | --- | --- |
| `show` | `bool` | `True` | 是否打印下载信息 |
| `timeout` | `float` | `None` | 超时时间（秒），`None`为无限等待 |
| `cancel_if_timeout` | `bool` | `False` | 如超时是否取消该任务 |

| 返回类型 | 说明 |
| --- | --- |
| `str` | 下载完成返回最终保存路径 |
| `False` | 超时或被取消返回`False` |

---

### 📌 取消任务[​](#-取消任务 "📌 取消任务的直接链接")

使用`DownloadMission`对象的`cancel()`方法，可取消任务。

调用该方法，已下载的文件会被删除，即使是已完成的任务。

[上一页

⤵️ download方法](/download/DownloadKit)[下一页

⚙️ 配置文件的使用](/advance/ini)

* [✅️ 概述](#️-概述)
  + [📌 功能](#-功能)
* [⚠️ 注意事项](#️-注意事项)
  + [📌 记得等待任务结束](#-记得等待任务结束)
  + [📌 多 Tab 操作时推荐设置临时路径](#-多-tab-操作时推荐设置临时路径)
  + [📌 启动下载管理功能](#-启动下载管理功能)
* [✅️ `click.to_download()`](#️-clickto_download)
* [✅️ 设置下载路径](#️-设置下载路径)
  + [📌 设置总下载路径](#-设置总下载路径)
  + [📌 设置 Tab 下载路径](#-设置-tab-下载路径)
* [✅️ 设置文件名](#️-设置文件名)
* [✅️ 等待](#️-等待)
  + [📌 等待下载开始](#-等待下载开始)
  + [📌 等待所有下载任务结束](#-等待所有下载任务结束)
  + [📌 等待某 Tab 所有下载任务结束](#-等待某-tab-所有下载任务结束)
* [✅️ 拦截下载任务](#️-拦截下载任务)
* [✅️ 同名文件的处理](#️-同名文件的处理)
  + [📌 自动重命名](#-自动重命名)
  + [📌 覆盖已有文件](#-覆盖已有文件)
  + [📌 跳过](#-跳过)
* [✅️ 任务管理](#️-任务管理)
  + [📌 获取任务信息](#-获取任务信息)
  + [📌 等待任务结束](#-等待任务结束)
  + [📌 取消任务](#-取消任务)

* 🧰 进阶使用
* ⬇️ 下载文件
* ⤵️ 概述

本页总览

⤵️ 概述
=====

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

DrissionPage 提供了强大的文件下载管理功能。

能够主动发起下载任务，也能够对浏览器触发的下载任务进行管理。

✅️️ `download()`方法[​](#️️-download方法 "️️-download方法的直接链接")
----------------------------------------------------------

该方法可以主动发起下载任务，提供任务管理、多线程、大文件分块、自动重连、文件名冲突处理等功能。

页面对象、`<iframe>`元素对象均支持此方法。

此方法是封装 requests 实现的，下载时会自动同步 cookies。

**示例：**

```
from DrissionPage import SessionPage  
  
page = SessionPage()  
page.download('https://dldir1.qq.com/qqfile/qq/TIM3.4.8/TIM3.4.8.22092.exe')
```

---

✅️️ 浏览器的下载任务[​](#️️-浏览器的下载任务 "✅️️ 浏览器的下载任务的直接链接")
-------------------------------------------------

浏览器页面对象、`<iframe>`对象可对浏览器下载任务进行控制。

包含以下功能：

* 每个标签页对象可独立指定下载地址
* 可在下载前指定重命名文件名
* 可拦截下载任务，获取任务信息

**示例：**

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
mission = tab('t:a').click.to_download('tmp', 'file_name')  # 点击一个会触发下载的链接，同时设置下载路径和文件名  
mission.wait()  # 等待下载结束
```

功能分解写法，效果和上面的一样：

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.set.download_path('save_path')  # 设置文件保存路径  
tab.set.download_file_name('file_name')  # 设置重命名文件名  
tab('t:a').click()  # 点击一个会触发下载的链接  
tab.wait.download_begin()  # 等待下载开始  
tab.wait.downloads_done()  # 等待下载结束
```

[上一页

🛩️ 启动配置](/SessionPage/session_opt)[下一页

⤵️ download方法](/download/DownloadKit)

* [✅️️ `download()`方法](#️️-download方法)
* [✅️️ 浏览器的下载任务](#️️-浏览器的下载任务)

* 🌟 特性演示
* ⭐ 下载文件

⭐ 下载文件
======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

DrissionPage 带一个简便易用的下载器，一行即可实现下载功能。

```
from DrissionPage import SessionPage  
  
url = 'https://www.baidu.com/img/flexible/logo/pc/result.png'  
save_path = r'C:\download'  
  
page = SessionPage()  
page.download(url, save_path)
```

[上一页

⭐ 获取元素属性](/features/features_demos/get_ele_attr)

## <a name="特性"></a>特性

* 💖 特性

本页总览

💖 特性
====

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

✅️️ 强大的自研内核[​](#️️-强大的自研内核 "✅️️ 强大的自研内核的直接链接")
----------------------------------------------

* 不依赖 webdriver
* 无需为不同版本的浏览器下载不同的驱动
* 运行速度更快
* 可以跨 iframe 查找元素，无需切入切出
* 可同时操作多个标签页，无需切换
* 方便好用的数据包监听功能
* 可处理非`open`状态的 shadow-root

---

✅️️ 无处不在的等待[​](#️️-无处不在的等待 "✅️️ 无处不在的等待的直接链接")
----------------------------------------------

网络环境不稳定，很多时候程序需要稍微等待一下才能继续运行。

等待太少，会导致程序出错，等待太多，又会浪费时间。

为了解决这些问题，本库在大量需要等待的部分内置了超时功能，并且可以随时灵活设置，大幅降低程序复杂性。

* 查找元素内置等待。可以为每次查找元素单独设定等待时间。使用灵活。
* 等待下拉列表选项。很多下拉列表使用 js 加载，本库选择下拉列表时，会自动等待列表项出现。
* 等待弹窗。有时预期的 alert 未必立刻出现，本库处理弹窗消息时也可以设置等待。
* 等待元素状态改变。使用`wait.ele()`方法可等待元素出现、消失、删除等状态。
* 等待页面进入加载状态或加载完成。不仅节省时间，还大幅提高程序稳定性。
* 点击功能也内置等待，如遇元素被遮挡可不断重试点击。
* 设置页面加载时限及加载策略。有时不需要完整加载页面资源，可根据实际需要设置加载策略。

---

✅️️ 极简的定位语法[​](#️️-极简的定位语法 "✅️️ 极简的定位语法的直接链接")
----------------------------------------------

本库制定了一套简洁高效的查找元素语法，支持链式操作，支持相对定位。

每次查找内置等待，可以独立设置每次查找超时时间。

同是设置了超时等待的查找，与 selenium 对比一下：

```
# 使用 selenium：  
element = WebDriverWait(driver, 10).until(ec.presence_of_element_located((By.XPATH, '//*[contains(text(), "some text")]')))  
  
# 使用 DrissionPage：  
element = tab('some text', timeout=10)
```

---

✅️️ 无需切入切出的使用方式[​](#️️-无需切入切出的使用方式 "✅️️ 无需切入切出的使用方式的直接链接")
----------------------------------------------------------

使用过 selenium 的人都知道，selenium 同一时间只能操作一个标签页或`<iframe>`。

需要用`switch_to()`方法来回切换，相当麻烦。

DrissionPage 则无需这些麻烦的操作，它把每个标签页和`<iframe>`都看作独立的对象，可以同时并发操作。

而且可以直接跨多层`<iframe>`获取里面的元素，然后直接处理，非常方便。

对比一下，获取 2 层`<iframe>`内一个 id 为`'div1'`的元素：

```
# 使用 selenium  
driver.switch_to.frame(0)  
driver.switch_to.frame(0)  
ele = driver.find_element(By.ID, 'div1')  
driver.switch_to.default_content()  
  
# 使用 DrissionPage  
ele = tab('#div1')
```

多标签页同时操作，selenium 无此功能：

```
tab1 = browser.get_tab(1)  
tab2 = browser.get_tab(2)  
  
tab1.get('https://www.baidu.com')  
tab2.get('https://www.163.com')
```

---

✅️️ 高度集成的便利功能[​](#️️-高度集成的便利功能 "✅️️ 高度集成的便利功能的直接链接")
----------------------------------------------------

很多操作方法集成了常用功能，如`click()`中内置`by_js`参数，可以直接改用 js 方式点击，而无需写独立的 js 语句。

数量太多，不一一阐述，可在使用中体验。

---

✅️️ 强大的下载功能[​](#️️-强大的下载功能 "✅️️ 强大的下载功能的直接链接")
----------------------------------------------

DrissionPage 内置一个下载工具，可实现大文件分块多线程下载文件。

并且可以直接读取缓存数据保存图片，而无需控制页面作另存操作。

---

✅️️ 自动重试连接[​](#️️-自动重试连接 "✅️️ 自动重试连接的直接链接")
-------------------------------------------

在访问网站时，由于网络不稳定可能导致连接异常。本库设置了连接自动重试功能，当网页连接异常，会默认重试 3 次。当然也可以手动设置次数和间隔。

```
tab.get('****', retry=5, interval=3)  # 出错时重试 5 次，每次间隔 3 秒
```

---

✅️️ 更多便捷的功能[​](#️️-更多便捷的功能 "✅️️ 更多便捷的功能的直接链接")
----------------------------------------------

* 可对整个网页截图，包括视口外的部分
* 每次运行程序可以反复使用已经打开的浏览器，无需每次从头运行
* s 模式访问网页时会自动纠正编码，无需手动设置
* s 模式在连接时会自动根据当前域名自动填写`Host`和`Referer`属性
* 下载工具支持多种方式处理文件名冲突、自动创建目标路径、断链重试等
* 支持直接获取`after`和`before`伪元素的内容
* 上传文件可直接拦截文件选择框并输入路径，无需依靠 GUI 或查找`<input>`元素输入

[上一页

💥 3.2 功能介绍](/features/3)[下一页

⭐ 与 requests 对比](/features/features_demos/requests)

* [✅️️ 强大的自研内核](#️️-强大的自研内核)
* [✅️️ 无处不在的等待](#️️-无处不在的等待)
* [✅️️ 极简的定位语法](#️️-极简的定位语法)
* [✅️️ 无需切入切出的使用方式](#️️-无需切入切出的使用方式)
* [✅️️ 高度集成的便利功能](#️️-高度集成的便利功能)
* [✅️️ 强大的下载功能](#️️-强大  的下载功能)
* [✅️️ 自动重试连接](#️️-自动重试连接)
* [✅️️ 更多便捷的功能](#️️-更多便捷的功能)

* 🔥 新版本介绍
* 💥 3.2 功能介绍

本页总览

💥 3.2 功能介绍
==========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

3.2 对比 3.1 有相当大的变化。既对底层逻辑进行了梳理，修复了许多问题，提高了稳定性，也对用户 api 进行了整合。

期待大佬们体验并提出意见建议。

沟通方式：

* QQ 群：897838127
* [点击提交 Issues](https://gitee.com/g1879/DrissionPage/issues)

✅️️ 特性改变[​](#️️-特性改变 "✅️️ 特性改变的直接链接")
-------------------------------------

### 📌 `WebPage`不再自动切换模式[​](#-webpage不再自动切换模式 "-webpage不再自动切换模式的直接链接")

旧版本中，调用非当前模式独有的方法，会先自动切换模式。比如在 d 模式调用`post()`方法，会自动切换到 s 模式。

3.2 中，只有显式调用`change_mode()`时才会切换模式。因此，可以在 s 模式时  控制浏览器，不会产生冲突。使用更灵活。

### 📌 `WebPage`的标签页对象也可以切换模式[​](#-webpage的标签页对象也可以切换模式 "-webpage的标签页对象也可以切换模式的直接链接")

`WebPage`的`get_tab()`方法现在返回`WebPageTab`对象，也能切换模式。创建时默认处于 d 模式。

### 📌 更改找不到元素时返回值[​](#-更改找不到元素时返回值 "📌 更改找不到元素时返回值的直接链接")

旧版本中，查找不到元素会返回`None`，3.2 版中会返回一个`NoneElement`对象。

该对象用`if`判断表现为`False`，调用其功能会抛出`ElementNotFoundError`异常。这样可以用`if`判断是否找到元素，也可以用`try`去捕获异常。

查找多个元素找不到时，依然返回空的`list`，这点和旧版本一致。

示例，当我们查找一个不存在的元素时：

```
ele = page.ele('****')
```

❌ 当前版本用`None`判断的方式将不可用：

```
if ele is None:  
    print('没有找到。')
```

⭕ 正确做法：

```
# 用if判断  
if not ele:  
    print('没有找到。')  
  
if ele:  
    print('找到了。')
```

```
# 用try捕获  
try:  
    ele.click()  
except ElementNotFoundError:  
    print('没有找到。')
```

### 📌 部分方法调整[​](#-部分方法调整 "📌 部分方法调整的直接链接")

* `run_js()`方法参数顺序作了调整
* `drag()`和`drag_to()`方法删除了`shake`参数
* 元素对象去除`wait_ele()`方法，改用`wait.****()`方法等待自身属性改变
* `ChromiumPage`的`to_tab()`、`close_tabs()`、`close_other_tabs()`方法可接收标签页对象
* `ActionChains`导入路径改为`from DrissionPage.common import ActionChains`
* `Keys`导入路径改为`from DrissionPage.common import Keys`

### 📌 下载默认使用浏览器[​](#-下载默认使用浏览器 "📌 下载默认使用浏览器的直接链接")

避免不了解 DownloadKit 的用户以为点击后程序没有反应。

现在点击后默认用浏览器下载到程序当前路径，可用`download_set.save_path()`设置保存位置。

✅️️ api 整合[​](#️️-api-整合 "✅️️ api 整合的直接链接")
-------------------------------------------

随着功能的增加，同类 api 也越来越多，比如`ChromiumPage`以`set_`开头的方法就有 9 个，元素点击的方式有 5 种。如何这些同类的 api 全堆在一起，会显得十分凌乱。

因此，3.2 版对大量 api 做了整合，避  免提示界面越来越臃肿，也增强了开发灵活性。

旧版的 api 目前还将保留，但将在以后的版本中舍弃。

旧版：

```
page.set_timeouts(20, 30, 40)
```

新版：

```
page.set.timeouts(20, 30, 40)
```

### 📌 浏览器页面对象[​](#-浏览器页面对象 "📌 浏览器页面对象的直接链接")

| 旧版 | 新版 | 说明 |
| --- | --- | --- |
| set\_timeouts() | set.timeouts() | 超时设置 |
| set\_session\_storage() | set.session\_storage() | session storage 设置 |
| set\_local\_storage() | set.local\_storage() | local storage 设置 |
| set\_user\_agent() | set.user\_agent() | user agent 设置 |
| set\_cookies() | set.cookies() | cookies 设置 |
| set\_headers() | set.headers() | headers 设置 |
| set\_page\_load\_strategy.\*\*\*\*() | set.load\_strategy.\*\*\*\*() | 页面加载策略设置 |
| set\_window.\*\*\*\*() | set.window.\*\*\*\*() | 浏览器大小位置设置 |
| set\_main\_tab() | set.main\_tab() | 主 tab 设置 |
| wait\_loading() | wait.load\_start() | 等待页面加载开始 |
| wait\_ele().\*\*\*\*() | wait.ele\_\*\*\*\*() | 等待元素变成某种状态 |
| scroll\_to\_see() | scroll.to\_see() | 把元素滚动到视口 |
| hide\_browser() | set.window.hide() | 隐藏浏览器 |
| show\_browser() | set.window.show() | 显示浏览器 |
| to\_front() | set.tab\_to\_front() | 设置某个标签页到激活状态 |

### 📌 元素对象[​](#-元素对象 "📌 元素对象的直接链接")

| 旧版 | 新版 | 说明 |
| --- | --- | --- |
| wait\_ele().\*\*\*\*() | wait.\*\*\*\*() | 等待元素变成某种状态 |
| r\_click() | click.right() | 右键点击元素 |
| m\_click() | clikc.middle() | 中键点击元素 |
| r\_click\_at() | click.right\_at() | 带偏移量中键点击元数据 |
| click\_at() | click.left\_at() | 带偏移量左键点击元素 |
| set\_attr() | set.attr() | 设置 attribute 属性 |
| set\_prop() | set.prop() | 设置 property 属性 |
| set\_innerHTML() | set.innerHTML() | 设置 innerHTML 内 容 |
| midpoint | locations.midpoint | 获取元素中点在页面位置 |
| client\_location | locations.viewport\_location | 获取元素左上角在视口位置 |
| client\_midpoint | locations.viewport\_midpoint | 获取元素中点在视口位置 |
| is\_selected | states.is\_selected | 获取元素是否被选中 |
| is\_displayed | states.is\_displayed | 获取元素是否显示 |
| is\_enabled | states.is\_enabled | 获取元素是否可操作 |
| is\_alive | states.is\_alive | 获取元素是否仍然在 DOM 内 |
| is\_in\_viewport | states.is\_in\_viewport | 获取元素是否在视口内 |
| pseudo\_before | pseudo.before | 获取 before 伪元素内容 |
| pseudo\_after | pseudo.after | 获取 after 伪元素内容 |
| obj\_id | ids.obj\_id | 获取元素 object id |
| node\_id | ids.node\_id | 获取元素 node id |
| backend\_id | ids.backend\_id | 获取元素 backend id |
| doc\_id | ids.doc\_id | 获取元素 doc id |

✅️️ 新增功能[​](#️️-新增功能 "✅️️ 新增功能的直接链接")
-------------------------------------

### 📌 拦截上传控件填写路径[​](#-拦截上传控件填写路径 "📌 拦截上传控件填写路径的直接链接")

之前的版本中，要上传文件需要开发者先在 DOM 内找到文件上传控件，有些经过伪装后实时加载的`<input>`元素并不好找，有时也会由 js 控制。

新版本中，再也无需费心查找上传控件，只要设置好要上传的路径，然后点击触发文本选择框，程序会自动拦截选择框，并把路径输入到控件，非常便利。

示例：

```
# 设置要上传的文件路径  
page.set.upload_files('demo.txt')  
# 点击触发文件选择框按钮  
btn_ele.click()  
# 等待路径填入  
page.wait.upload_paths_inputted()
```

点击按钮后，文本选择框被拦截不会弹出，但可以看到文件路径已经传入其中。

由于此动作是异步输入，需显式等待输入完成才进行下一步操作。

### 📌 优先读取项目路径 ini 文件[​](#-优先读取项目路径-ini-文件 "📌 优先读取项目路径 ini 文件的直接链接")

旧版本中默认 ini 文件存放在 DrissionPage 安装目录下，修改要通过代码进行，给调试带来不便。新版会优先在用户项目文件夹下查找`'dp_configs.ini'`文件并使用，使开发时可方便地手动更改配置。项目打包也可以直接打包而不会造成找不到文件问题。

`easy_set`方法中增加`configs_to_here()`方法，调用该方法可直接把默认 ini 文件复制到当前路径。

### 📌 查找元素增加或语法[​](#-查找元素增加或语法 "📌 查找元素增加或语法的直接链接")

查找元素增加`@|`语法，用于或关系匹配多个属性：

```
page('@|class=xxx@|name=abc')
```

或语法`@|`不能跟与语法`@@`共同生效。

### 📌 找不到元素时可抛出异常[​](#-找不到元素时可抛出异常 "📌 找不到元素时可抛出异常的直接链接")

找不到元素时，除了返回`NoneElement`对象，也可设置直接抛出异常。

```
from DrissionPage.easy_set import raise_when_ele_not_found  
  
raise_when_ele_not_found(True)
```

该设置是全局设置，设置后整个项目都会生效。

**示例：**

```
from DrissionPage import SessionPage  
from DrissionPage.easy_set import raise_when_ele_not_found  
  
raise_when_ele_not_found(True)  
  
page = SessionPage()  
page.get('https://www.baidu.com')  
ele = page('****')
```

上面的代码会抛出`ElementNotFound`异常。

### 📌 新增一批异常[​](#-新增一批异常 "📌 新增一批异常的直接链接")

新增一批异常，调用位置：

```
from DrissionPage.errors import ElementLossError
```

* `AlertExistsError`：调用页面功能若存在未处理的弹出框则抛出
* `ContextLossError`：页面被刷新后仍调用其中的元素时抛出
* `ElementLossError`：元素因页面或自身被刷新而失效时抛出
* `CallMethodError`：调用 cdp 时产生的异常
* `TabClosedError`：标签页关闭后仍调用其功能时抛出
* `ElementNotFoundError`：找不到元素抛出
* `JavaScriptError`：JavaScript 运行错误
* `NoRectError`：对没有大小和位置信息的元素获取这些信息时抛出

### 📌 新增的方法和属性[​](#-新增的方法和属性 "📌 新增的方法和属性的直接链接")

`ChromiumPage`：

| 名称 | 说明 |
| --- | --- |
| run\_cdp\_loaded() | 执行 cdp 命令，执行前等待页面加载完成 |
| run\_js\_loaded() | 执行 js 语句，执行前等待页面加载完成 |
| wait.load\_complete() | 等待页面加载完毕 |
| wait.upload\_paths\_inputted() | 等待上传文件路径输入到文件选择框 |
| rect.browser\_location | 获取浏览器左上角在屏幕上坐标 |
| rect.page\_location | 获取页面左上角在屏幕上坐标 |
| rect.viewport\_location | 获取视口在屏幕上坐标 |
| rect.browser\_size | 获取浏览器大小 |
| rect.page\_size | 获取页面大小 |
| rect.viewport\_size | 获取视口大小，不含滚动条 |
| rect.viewport\_size\_with\_scrollbar | 获取视口大小，含滚动条 |
| remove\_ele() | 从页面上移除一个元素 |
| get\_frame() | 获取一个`ChromiumFrame`对象 |

`SessionPage`新增一系列设置方法：

| 名称 | 说明 |
| --- | --- |
| set.header() | 设置一个 header 值 |
| set.proxies() | 设置代理 |
| set.auth() | 设置登录信息 |
| set.hooks() | 设置回调方法 |
| set.params() | 设置连接参数 |
| set.cert() | 设置证书 |
| set.stream() | 设置是否使用流式响应内容 |
| set.trust\_env() | 设置是否信任环境 |
| set.max\_redirects() | 设置最大重定向次数 |
| set.max\_redirects() | 添加适配器 |

`ChromiumElement`：

| 名称 | 说明 |
| --- | --- |
| locations.screen\_location | 获取元素左上角在屏幕上的坐标 |
| locations.screen\_midpoint | 获取元素中间点在屏幕上的坐标 |
| locations.screen\_click\_point | 获取元素点击点在屏幕上的坐标 |
| locations.click\_point | 获取元素点击点在页面上的坐标 |
| locations.viewport\_click\_point | 获取元素点击点在视口上的坐标 |
| states.is\_covered | 获取元素是否被覆盖 |
| click.at() | 带偏移量点击元素，可指定按键 |
| wait.covered() | 等待元素被覆盖 |
| wait.covered() | 等待元素不被覆盖 |

### 📌 命令行工具[​](#-命令行工具 "📌 命令行工具的直接链接")

新增支持命令行工具。

* `--set-browser-path`（`-p`）：设置配置文件中的浏览器路径
* `--set-user-path`（`-u`）：设置配置文件中的用户数据路径
* `--configs-to-here`（`-c`）：复制默认配置文件到当前路径
* `--launch-browser`（`-l`）：启动浏览器，传入端口号，0表示用配置文件中的值

```
dp --set-browser-path '/Application/Goolge Chrome.app/Contents/MacOS/Google Chrome'
```

✅️️ 优化和问题修复[​](#️️-优化和问题修复 "✅️️ 优化和问题修复的直接链接")
----------------------------------------------

* 对程序底层和业务逻辑进行了重新梳理，优化程序逻辑，大幅增强稳定性
* 新旧版本完全隔离，新版以后开发可放飞自我，无需担心影响以前用`MixPage`开发的程序
* 现在会返回开发者能看懂的异常信息
* 修复页面加载和退出触发弹窗引起的问题
* 修复`<iframe>`加载时可能出现的 500 错误
* 修复异域`<iframe>`点击问题
* 没有位置和大小信息的元素在获取这些信息时，现在会抛出异常
* 修复内存没有正确释放的问题
* 修复点击被固定栏遮挡问题
* 接管新出现的`<iframe>`会自动等待内容加载
* 修复`<iframe>`在同域和异域间互相跳转时会卡住的问题
* 修复`<iframe>`内元素截  图出现偏移问题

[上一页

💥 4.0 功能介绍](/features/4)[下一页

💖 特性](/features/)

* [✅️️ 特性改变](#️️-特性改变)
  + [📌 `WebPage`不再自动切换模式](#-webpage不再自动切换模式)
  + [📌 `WebPage`的标签页对象也可以切换模式](#-webpage的标签页对象也可以切换模式)
  + [📌 更改找不到元素时返回值](#-更改找不到元素时返回值)
  + [📌 部分方法调整](#-部分方法调整)
  + [📌 下载默认使用浏览器](#-下载默认使用浏览器)
* [✅️️ api 整合](#️️-api-整合)
  + [📌 浏览器页面对象](#-浏览器页面对象)
  + [📌 元素对象](#-元素对象)
* [✅️️ 新增功能](#️️-新增功能)
  + [📌 拦截上传控件填写路径](#-拦截上传控件填写路径)
  + [📌 优先读取项目路径 ini 文件](#-优先读取项目路径-ini-文件)
  + [📌 查找元素增加或语法](#-查找元素增加或语法)
  + [📌 找不到元素时可抛出异常](#-找不到元素时可抛出异常)
  + [📌 新增一批异常](#-新增一批异常)
  + [📌 新增的方法和属性](#-新增的方法和属性)
  + [📌 命令行工具](#-命令行工具)
* [✅️️ 优化和问题修复](#️️-优化和问题修复)

* 🔥 新版本介绍
* 💥 4.1 功能介绍

本页总览

💥 4.1 功能介绍
==========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

4.1 的最主要变化在于用`Chromium`对象替代`ChromiumPage`和`WebPage`对象。

从 DrissionPage 诞生的时候开始，就存在`DriverPage`和`MixPage`对象（后来的`ChromiumPage`和`WebPage`）用于控制浏览器。

这是因为 DrissionPage 最开始是基于 selenium 开发的。

selenium 只能通过切换焦点的方式操作标签页，无法同时操作多个。

所以一个 Page 对象对应一个 Driver 对象就足够了。

但随着新版的开发，基于 selenium 的设计逐渐被替换，Page 对象作为最后的遗留，不足也日渐显现：

* 它是浏览器和一个标签页的合体，在概念上不容易理解
* 它在管理标签页的开关、查找时候容易造成混淆
* 它代码结构臃肿，不利于后续功能开发

而它的唯一优点，是在创建的时候可以少写一行代码。

因此，在 4.1，会刻意淡化 Page 对象的存在感，而改用`Chromium`对象作为程序入口。

由于已有大量项目使用`ChromiumPage`和`WebPage`进行开发，这两个对象仍然保留，功能不会有太大变化。

✅️ `Chromium`对象[​](#️-chromium对象 "️-chromium对象的直接链接")
-----------------------------------------------------

`Chromium`对象用于连接浏览器，管理标签页和全局运行参数。

### 📌 连接浏览器[​](#-连接浏览器 "📌 连接浏览器的直接链接")

```
from DrissionPage import Chromium  
  
browser = Chromium()
```

---

### 📌 标签页操作[​](#-标签页操作 "📌 标签页操作的直接链接")

```
from DrissionPage import Chromium  
  
browser = Chromium()  
tab1 = browser.latest_tab  # 获取最后激活的标签页对象  
tab2 = browser.get_tab(title='DrissionPage')  # 获取指定标签页  
tab3 = browser.new_tab()  # 新建标签页  
browser.activate_tab(tab2)  # 将tab2提到最前面  
tab1.close()  # 关闭标签页
```

---

### 📌 浏览器操作[​](#-浏览器操作 "📌 浏览器操作的直接链接")

```
from DrissionPage import Chromium  
  
browser = Chromium()  
browser.set.cookies({'abc': '123'})  # 设置cookies  
browser.set.download_path('C:\\tmp')  # 设置下载路径  
# 更多详见相关章节
```

---

✅️ api 变化[​](#️-api-变化 "✅️ api 变化的直接链接")
----------------------------------------

* `WebPageTab`改名为`MixTab`
* `SessionPage`、`ChromiumPage`和`WebPage`初始化时删除`timeout`提示，以后会废弃
* `activate_tab()`取代`set.tab_to_front()`
* 所有对象增加`find()`方法，用于同时匹配多个定位符
* 页面对象增加`console`属性，可读取控制台信息
* Frame 对象增加`set.property()`、`set.style()`、`link`
* Tab 对象的`close()`方法增加`others`参数
* `quit()`增加`del_data`参数
* `cookies()`删除`as_dict`参数，增加`as_dict()`、`as_json`和`as_str()`方法
* 浏览器页面和元素对象的`s_ele()`和`s_eles()`方法增加`tiemout`参数
* 浏览器页面和元素对象增加`rect.scroll_position`属性
* 元素对象增加`get_frame()`方法
* 元素对象增加`timeout`属性
* `parent()`和 shadow-root 内查找方法增加`timeout`参数
* 动作链删除`db_click()`，各点击方法增加`times`参数
* `wait.new_tab()`增加`curr_tab`参数
* 滚动增加`scroll()`方法
* `ChromiumOptions`增加`new_env()`方法，ini 文件增加`new_env`参数，用于指定必须用新环境
* `ChromiumOptions`增加`is_headless`属性
* `auto_port()`方法删除`tmp_path`参数
* `wait.alert_closed()`增加`timeout`参数

---

✅️ 行为变化[​](#️-行为变化 "✅️ 行为变化的直接链接")
----------------------------------

* `Chromium`只返回`MixTab`类型的标签页对象
* `ChromiumFrame`对象默认改为单例
* `MixTab`和`MixPage`的`post()`方法必返回`Response`对象
* 部分等待方法会返回调用者，方便链式操作
* 元素对象各种动作返回元素本身，便于链式操作
* 打印`NoneElement`改成详细信息
* `src()`方法可获取`<link>`指向的文件内容
* 录像改为 H.265 编码
* `shadow_root`属性增加等待附加到元素（超时 10 秒）
* `set.cookies()`忽略过期 cookie
* `timeout`属性不再接受赋值

---

✅️ 优化和问题 修复[​](#️-优化和问题修复 "✅️ 优化和问题修复的直接链接")
--------------------------------------------

* 优化连接浏览器失败报错
* 优化`css_path`
* 修复`new_tab()`在访客模式和隐私模式的问题
* 修复 Frame 对象滚动填入`tuple`定位符报错问题
* 修复`states.is_displayed`有些情况下不正确问题
* 修复元素`link`属性不正确的问题
* 修复 shadow-root 内用 css 找元素的一个问题
* 修复异域`<iframe>`内元素屏幕坐标不准问题
* 修复`new_tab=True`时下载路径不正确问题
* 修复`attr()`填入大写字母无法获取问题

[下一页

💥 4.0 功能介绍](/features/4)

* [✅️ `Chromium`对象](#️-chromium对象)
  + [📌 连接浏览器](#-连接浏览器)
  + [📌 标签页操作](#-标签页操作)
  + [📌 浏览器操作](#-浏览器操作)
* [✅️ api 变化](#️-api-变化)
* [✅️ 行为变化](#️-行为变化)
* [✅️ 优化和问题修复](#️-优化和问题修复)

* 🔥 新版本介绍
* 💥 4.0 功能介绍

本页总览

💥 4.0 功能介绍
==========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

3.x 是自主研发底层的初步尝试，很多地方摸着石头过河，存在一些不太成熟的地方。

经过一段时间的使用，积累一些经验之后，4.0 在 3.x 的基础上对底层进行了大幅重构，新增大量功能，改善运行效率和稳定性，优化项目结构，解决很多存在的问题。对比旧版本有质的提高。

但同时不少 api 发生了变化，不能完全兼容旧版本。

api 的变化有些是功能优化必需的改变，有些则是本人对命名简洁的执念，趁着大版本的更新顺便把长期不太满意的命名给改了。

给使用者造成一定不便感到抱歉，但长痛不如短痛，趁着项目用的人不多，干脆舍弃历史包袱果断改掉。

有些原来的写法在 4.0.0 中还能正常使用，但 IDE 会提示无效，将在以后的版本中完全删除。推荐尽快更新为新写法。

本节仅简述功能变化，具体使用方法详见各对应章节。

✅️ 新的抓包功能[​](#️-新的抓包功能 "✅️ 新的抓包功能的直接链接")
----------------------------------------

3.2 中，抓包功能主要由 FlowViewer 和`wait.data_packets()`提供。

FlowViewer 是本人的一个练手作品，写得比较随意，技术也还没到家。存在漏抓、信息不全、api 不够合理的问题。

4.0 中，每个页面对象都内置了监听器，能力全面升级，api 也更合理。

### 📌 旧 api 变化[​](#-旧-api-变化 "📌 旧 api 变化的直接链接")

* 弃用 FlowViewer，以后也不会再升级
* 删除`wait.set_targets()`
* 删除`wait.stop_listening()`方法
* 删除`wait.data_packets()`方法
* `DrissionPage.common`路径删除`FlowViewer`

---

### 📌 新 api[​](#-新-api "📌 新 api的直接链接")

* 每个标签页对象（包括`ChromiumFrame`）新增`listen`属性，内置监听功能
* 用`listen.start()`和`listen.stop()`启动和停止监听
* 用`listen.wait()`阻塞等待数据包
* 用`listen.steps()`同步获取监听结果
* 增加`listen.wait_silent()`等待所有请求完成（包含 targets 以外的）
* 监听结果结构优化，request 和 response 数据分开存放

---

### 📌 示例[​](#-示例 "📌 示例的直接链接")

下面示例可直接运行查看结果。这个示例会计时，用于与下个示例对比。

```
from DrissionPage import ChromiumPage  
from TimePinner import Pinner  
from pprint import pprint  
  
page = ChromiumPage()  
page.listen.start('api/getkeydata')  # 指定监听目标并启动监听  
pinner = Pinner(True, False)  
page.get('http://www.hao123.com/')  # 访问网站  
packet = page.listen.wait()  # 等待数据包  
pprint(packet.response.body)  # 打印数据包正文  
pinner.pin('用时', True)
```

**输出：**

```
{'hao123.new.shishi.bangdan.recom': [{'index': '1',  
                                      'pure_title': '以色列和哈马斯移交首批被扣押人员'},  
                                     {'index': '2',  
                                      'pure_title': '听到免签政策法国外长笑了'},  
                                     ......  
用时：3.3114853000151925
```

---

✅️ 新的页面访问逻辑[​](#️-新的页面访问逻辑 "✅️ 新的页面访问逻辑的直接链接")
----------------------------------------------

3.x 中存连接存在以下主要问题：

* 浏览器页面对象`get()`方法的`timeout`参数只对加载阶段生效，无法覆盖连接阶段；
* 加载策略`none`模式没有实际用处。

这两个问题都在 4.0 中解决，且能够让用户自主控制终止连接的时机。
另外还对连接逻辑进行了优化，避免卡死情况出现。

### 📌 api 变化[​](#-api-变化 "📌 api 变化的直接链接")

* 页面对象`page_load_strategy`属性改名为`load_mode`
* `set.load_strategy`改为`set.load_mode`

---

### 📌 行为变化[​](#-行为变化 "📌 行为变化的直接链接")

* `get()`方法的`timeout`参数现在可覆盖整个过程
* `timeout`参数对非`get()`方法触发的加载（如点击链接）也能生效
* `SessionPage`和`WebPage`的 s 模式，如收到空数据，也会重试
* `SessionPage`的`get()`方法可以指向本地文件

---

### 📌 新的`none`加载模式[​](#-新的none加载模式 "-新的none加载模式的直接链接")

旧版中，`none`加载策略是当页面  连接成功就立刻停止加载，这在实际使用时没有什么意义。

新版中，这个模式改成：除非加载完成，否则程序不会主动将其停止（即使已超时），同时连接状态不再阻塞程序，而允许用户进行状态判断，主动停止加载。

这样提供给用户非常大的自由度，可等到关键数据包或元素出现就主动停止页面加载，大幅提升执行效率。

---

### 📌 示例[​](#-示例-1 "📌 示例的直接链接")

我们继续使用上一个示例的代码，但把加载模式设为`none`，且获取到数据时主动停止加载。

```
from DrissionPage import ChromiumPage  
from TimePinner import Pinner  
from pprint import pprint  
  
page = ChromiumPage()  
page.set.load_mode.none()  # 设置加载模式为none  
page.listen.start('api/getkeydata')  # 指定监听目标并启 动监听  
pinner = Pinner(True, False)  
page.get('http://www.hao123.com/')  # 访问网站  
packet = page.listen.wait()  # 等待数据包  
page.stop_loading()  # 主动停止加载  
pprint(packet.response.body)  # 打印数据包正文  
pinner.pin('用时', True)
```

**输出：**

```
{'hao123.new.shishi.bangdan.recom': [{'index': '1',  
                                      'pure_title': '以色列和哈马斯移交首批被扣押人员'},  
                                     {'index': '2',  
                                      'pure_title': '听到免签政策法国外长笑了'},  
                                     ......  
用时：1.2575092000188306
```

可见节省了2秒时间。
当网站要访问一些不稳定资源时，节省的时间相当客观，也能提高程序的稳定性。

---

✅️ 新的下载管理功能[​](#️-新的下载管理功能 "✅️ 新的下载管理功能的直接链接")
----------------------------------------------

在旧版中，下载管理功能存在以下问题：

* 浏览器下载管理和内置下载器`DownloadKit`的配置都使用`download_set`属性进行设置，容易造成混淆。
* 浏览器下载任务不能在下载前指定文件名
* 该功能有随着浏览器版本更新失效的风险

4.0 对浏览器下载管理功能进行了完完全全的重构，结构更为合理，功能更多。
同时，内置下载器的设置和浏览器下载任务设置进行了分离。

### 📌 api 变化[​](#-api-变化-1 "📌 api 变化的直接链接")

* 页面对象删除`download_set`属性
* 增加`set.download_path()`方法
* 增加`set.download_file_name()`方法

---

### 📌 新增功能[​](#-新增功能 "📌 新增功能的直接链接")

* Tab 对象和 Frame 对象也支持`download()`方法
* 每个 Tab 对象可单独设置下载路径和重命名文件名
* 可拦截浏览器下载任务并获取其信息
* 可取消浏览器下载任务、获取下载进度、等待任务完成
* 可设置遇到文件夹已存在时的处理方式

---

### 📌 行为变化[​](#-行为变化-1 "📌 行为变化的直接链接")

4.0 中默认不启用浏览器下载任务管理，只有在启动参数中设置了下载路径，或调用`set.download_path()`方法时才会启动。

未启动任务管理功能时，下载行为和普通使用一样。

---

### 📌 示例[​](#-示例-2 "📌 示例的直接链接")

以下示例可直接运行。

```
from DrissionPage import ChromiumPage  
  
page = ChromiumPage()  
page.get('https://office.qq.com/download.html')  
page.set.download_path('tmp')  # 设置文件保存路径  
page.set.download_file_name('qq')  # 设置文件名  
page('#downloadWin').click()  # 点击触发下载  
mission = page.wait.download_begin()  # 等待下载开始并获取任务对象  
mission.wait()  # 等下下载任务完成
```

**输出：**

```
url：https://dldir1.qq.com/qqfile/qq/TIM3.4.8/TIM3.4.8.22124.exe  
文件名：qq.exe  
目标路径：D:\coding\projects\DrissionPage\tmp  
100.0% 下载完成 D:\coding\projects\DrissionPage\tmp\qq.exe
```

---

✅️ 页面对象[​](#️-页面对象 "✅️ 页面对象的直接链接")
----------------------------------

这里说的页面对象包括 Page 对象（`ChromiumPage`、`WebPage`）、Tab 对象（`ChromiumTab`、`WebPageTab`）、`ChromiumFrame`对象。

### 📌 启动参数变化[​](#-启动参数变化 "📌 启动参数变化的直接链接")

4.0 中，创建`WebPage`和`ChromiumPage`对象时，不再接收`ChromiumDriver`对象。

意思是不再支持传递控制权的方式创建页面对象。

因为本身支持多个页面对象控制同一个标签页，如果需要多页面对象协同，只要`get_tab()`创建一个新对象就行，可和原有对象并行使用。

而且，传递控制权本身有稳定性方面隐患，因此新版中将其删除。

相应地，启动参数的名称也发生了变化：

* `WebPage`对象的`driver_or_options`参数改名为`chromium_options`
* `ChromiumPage`对象的`addr_driver_opts`参数改名为`addr_or_opts`

另外，`ChromiumPage`的启动参数`addr_or_opts`现在可以接收`int`数据，直接传入端口号。

---

### 📌 内置动作链[​](#-内置动作链 "📌 内置动作链的直接链接")

4.0 中，每个页面对象内置`actions`属性，即动作链。

内置的动作链与直接创建的动作链对象有一个不同点，每次操作会等待页面加载完成再执行。

**示例：**

```
page.actions.hold(ele).move(50).release()
```

---

### 📌 状态信息[​](#-状态信息 "📌 状态信息的直接链接")

旧版中，页面对象拥有`ready_state`、`is_loading`、`is_alive`属性，现在都合并到`states`属性中。

```
# ------ 旧版代码 ------  
print(page.is_loading)  
  
# ------ 新版代码 ------  
print(page.states.is_loading)
```

---

### 📌 其它[​](#-其它 "📌 其它的直接链接")

* `ChromiumPage`和`WebPage`改为固定单例
* `get_tab()`获取的 Tab 对象默认单例，可用`Settings`设置允许多例
* 页面对象增加`raw_data`参数，s 模式下返回原始数据
* 所有页面对象增加`close()`方法，`SessionPage`用于关闭连接，浏览器页面对象用于关闭标签页
* 浏览器页面对象增加`wait()`方法，用于等待若干秒
* 浏览器页面对象增加`wait.ele_loaded()`方法，等待元素加载到DOM
* 浏览器页面对象增加`wait.title_change()`和`wait.url_change()`方法，用于等待 title 和 url 变化
* 浏览器页面对象增加`wait.alert_closed()`方法，用于等待弹窗被手动关闭
* 浏览器页面对象增加`set.cookie()`方法，可设置单个 cookie
* 浏览器页面对象增加`set.blocked_urls()`方法，可设置忽略的连接
* Tab 和 Page 对象增加`disconnect()`方法，用于断开与网页连接
* Tab 和 Page 对象增加`reconnect()`方法，用于断开并重新连接网页
* Tab 和 Page 对象增加`save()`方法，用于把网页保存为 mhtml
* Tab 和 Page 对象增加`add_init_js()`和`remove_init_js()`方法
* `quit()`方法增加`force`参数，可强制关闭浏览器进程
* `ChromiumFrame`增加`ract`属性
* `ChromiumFrame`的`frame_size`属性改为`rect.size`
* `wait.ele_delete()`方法改为`wait.ele_deleted()`
* `wait.ele_display()`方法改为`wait.ele_displayed()`
* `wait.load_complete()`方法改为`wait.doc_loaded()`
* 优化`SessionPage`和`WebPage`s 模式访问速度
* `WebPage`在 d 模式时，`post()`返回`Response`对象

---

✅️ cookies 设置[​](#️-cookies-设置 "✅️ cookies 设置的直接链接")
----------------------------------------------------

* `set.cookies()`可接收单个 cookie
* 增加`set.cookies.clear()`方法用于清除 cookies
* 增加`set.cookies.remove()`方法用于删除一个 cookie 项

---

✅️ 标签页管理[​](#️-标签页管理 "✅️ 标签页管理的直接链接")
-------------------------------------

### 📌 不再支持`to_tab()`功能[​](#-不再支持to_tab功能 "-不再支持to_tab功能的直接链接")

`to_tab()`的设计源自于 selenium，用于在多个标签页间切换程序焦点。

selenium 没有 tab 对象，driver 每次只能操作一个 tab。多 tab 使用时需在不同的 tab 间来回切换，且切换的时候会丢失之前获取过的元素，效率低，使用不便。

DrissionPage 3.x 开始就支持多 tab 对象共存，对象之间互不影响，而且标签页无需激活即可操作。因此不再需要切换标签页。

而且焦点切换时如果页面正在加载，实现逻辑较为复杂，也会有稳定性问题。

基于此，决定删除`to_tab()`方法，使用`get_tab()`取代之。

**涉及的 api修改：**

* 删除`to_tab()`方法
* 删除`to_main_tab()`、`set.main_tab()`方法
* 删除`main_tab`属性
* `new_tab()`方法删除`switch_to`参数

**新建标签页并切换到新标签页：**

```
# ------ 旧版代码 ------  
tab = page.new_tab(switch_to=True)  
  
# ------ 新版代码 ------  
tab = page.new_tab()
```

**操作另一个标签页**

```
# ------ 旧版代码 ------  
page.to_tab(page.tabs[1])  
  
# ------ 新版代码 ------  
tab = page.get_tab(1)  # 创建一个可与page同时使用的tab对象
```

---

### 📌 `new_tab()`的新功能[​](#-new_tab的新功能 "-new_tab的新功能的直接链接")

Page 对象的`new_tab()`方法增加了 3 个参数：

* `new_window`：是否创建新的窗口，新窗口与旧标签页同属一个浏览器，只是独立窗口
* `background`：新建的标签页是否为不激活状态（即使不激活也可以操作）
* `new_context`：是否创建独立的隐身窗口，该窗口与旧标签页 cookies 相互独立

现在`new_tab()`返回新建的标签页对象，而非其 tab\_id。

**示例：**

```
tab = page.new_tab()  
tab.get('http://DrissionPage.cn')
```

---

### 📌 标签页的位置与大小[​](#-标签页的位置与大小 "📌 标签页的位置与大小的直接链接")

在旧版中，只有`ChromiumPage`或`WebPage`对象可以设置窗口位置、大小、状态。

Tab 对象（`ChromiumTan`、`WebPageTab`）虽然能获取窗口窗口上述信息，但只是获取 Page 所控制的标签页信息。

在 4.0 中，独立窗口的标签页，也能设置和获取上面这些属性。

以下属性和方法名称进行了修改:

* `rect.borwser_size`改为`rect.window_size`
* `rect.borwser_location`改为`rect.window_location`
* `set.window.maximized()`改为`set.window.max()`
* `set.window.minimized()`改为`set.window.mini()`
* `set.window.fullscreen()`改为`set.window.full()`

**示例：**

```
tab = page.get_tab(1)  
print(tab.rect.window_state)  # 获取窗口状态  
print(tab.rect.window_location)  # 获取窗口位置  
print(tab.rect.window_size)  # 获取窗口大小  
  
tab.set.window.size(500, 500)  # 设置窗口大小  
tab.set.window.location(500, 500)  # 设置窗口位置  
tab.set.window.max()  # 窗口最大化  
# 更多详见相关文档……
```

---

### 📌 标签页的行为[​](#-标签页的行为 "📌 标签页的行为的直接链接")

在旧版中，标签页的开关、激活由 Page 对象管理。JS 弹窗也只有 Page 对象能处理。

在 4.0 中，标签页对象可以激活、关闭自己，也可以处理自己的弹出对话框。

* `tab.set.activate()`：标签页对象激活自己
* `tab.close()`：标签页对象关闭自己
* `tab.handle_alert()`：标签页对象处理自己的弹窗
* `tab.states.has_alert`：标签页增加此属性表示是否有弹窗存在

---

### 📌 `get_tab()`参数变化[​](#-get_tab参数变化 "-get_tab参数变化的直接链接")

现在`get_tab()`方法可接收标签页序号（序号从 0 开始）。`tab_id`参数改为`id_or_num`。

但需要注意，序号与标签页视觉排序不一定一致，而是按照激活顺序排列。

**示例：**

```
tab = page.get_tab(1)  # 获取列表中第二个标签页的对象
```

---

✅️ 元素相关[​](#️-元素相关 "✅️ 元素相关的直接链接")
----------------------------------

### 📌 定位语法变化[​](#-定位语法变化 "📌 定位语法变化的直接链接")

旧版中，定位语法中使用`@@-`或`@|-`表示否定，但视觉效果不明显，意义也不太明确。

因此改用`@!`替代，使其更能体现否定的含义。

`@!`可与`@@`或`@|`混用，与还是或关系视`@@`还是`@|`而定。

也可以单独使用，否定某个单独属性。

**示例：**

```
# ------ 旧版语法 ------  
page.ele('@@arg1=abc@@-arg2=def')  
# ------ 新版语法 ------  
page.ele('@@arg1=abc@!arg2=def')  
  
# ------ 旧版语法 ------  
page.ele('t:div@|arg1=abc@|-arg2=def')  
# ------ 新版语法 ------  
page.ele('t:div@|arg1=abc@!arg2=def')  
  
# ------ 旧版语法 ------  
page.ele('@@-arg1=abc')  
# ------ 新版语法 ------  
page.ele('@!arg1=abc')
```

---

### 📌 相对定位参数优化[​](#-相对定位参数优化 "📌 相对定位参数优化的直接链接")

旧版本中，相对定位如果想用序号获取元素，需要写成`ele.next(index=1)`。

但是我想 把这个语句变得更简化一点，写成`ele.next(1)`即可表示获取下一个兄弟元素。

4.0 中支持这种写法，当`filter_loc`参数接收`int`类型参数，即可当作`index`参数使用。

`parent()`方法增加`index`参数，当`level_or_loc`传入定位符，使用此参数选择第几个结果。

---

### 📌 位置和大小[​](#-位置和大小 "📌 位置和大小的直接链接")

旧版中，元素大小和位置信息由`location`、`locations`、`size`几个属性提供。

这些属性都与形状相关，为使逻辑更清晰，与页面对象逻辑一致，将它们统一归纳到`rect`属性中。

* 删除`size`、`location`、`locations`属性，新增`rect`属性
* 旧版中`loactions.****`的属性改为`rect.****`
* 大小和位置信息，从`int`类型改为`float`类型
* 增加`states.has_rect`属性，返回元素是否拥有大小和位置
* 增加`states.is_whole_in_viewport`属性，返回元素是否整个都在视口内

```
# ------ 旧版代码 ------  
ele.size  
ele.location  
ele.locations.midpoint  
  
# ------ 新版代码 ------  
ele.rect.size  
ele.rect.location  
ele.rect.midpoint
```

---

### 📌 点击[​](#-点击 "📌 点击的直接链接")

* `click()`增加`wait_stop`参数，默认等待元素运动停止再点击
* `click()`默认等待元素运动停止再执行点击
* `click.twice()`改为`click.multiple()`

---

### 📌 查找元素失败显示细节[​](#-查找元素失败显示细节 "📌 查找元素失败显示细节的直接链接")

旧版链式查找元素时，遇到其中一个查找失败，不能很直观地显示哪个语句失败。

在 4.0 中，可以把查找失败的语句和定位语句显示在报错信息中。

**示例**

```
from DrissionPage import ChromiumPage  
  
page = ChromiumPage(timeout=1)  
page.get('https://baidu.com')  
print(page('#wrapper')('#s_tab')('#abcd').text)  # ('#abcd')这个元素不存在
```

输出：

```
DrissionPage.errors.ElementNotFoundError:   
没有找到元素。  
method: ele()  
args: {'locator': '#abcd'}
```

---

### 📌 设置查找失败元素返回默认值[​](#-设置查找失败元素返回默认值 "📌 设置查找失败元素返回默认值的直接链接")

如果查找元素后要获取一个属性，但这个元素不一定存在，或者链式查找其中一个节点找不到，可以设置查找失败时返回的值，而不是抛出异常。

这样可以简化一些采集逻辑。

**示例**

比如说，遍历页面上一个列表中多个对象，但其中有些元素可能缺失某个子元素，旧版中要这样写：

```
from DrissionPage import ChromiumPage  
  
page = ChromiumPage()  
for li in page.eles('t:li'):  
    ele = li('.name')  
    name = ele.text if ele else None  
    ele = li('.age')  
    age = ele.text if ele else None  
    ele = li('.phone')  
    phone = ele.text if ele else None
```

在新版中，可以这样写：

```
from DrissionPage import ChromiumPage  
  
page = ChromiumPage()  
page.set.NoneElement_value('没找到')  
for li in page.eles('t:li'):  
    name = li('.name').text  
    age = li('.age').text  
    phone = li('.phone').text
```

这样，假如某个子元素不存在，不会抛出异常，而是返回`'没找到'`这个字符串。

---

### 📌 更多[​](#-更多 "📌 更多的直接链接")

* `ele()`和`s_ele()`增加`index`参数，可指定获取第几个
* 增加`wait.stop_moving()`方法，可等待移动结束
* 增加`wait()`方法，用于等待若干秒
* 增加`check()`方法，可选中或取消选中元素
* 增加`wait.has_rect()`方法，用于等待元素拥有大小和位置
* 滚动添加`to_center()`方式，可滚动到视口中央
* 增加`select.by_option()`和`select.cancel_by_option()`方法，可选取列表项元素
* 增加`states.has_rect`属性
* 元素被覆盖时，`states.is_covered`属性返回覆盖元素的 id
* 添加`states.is_whole_in_viewport`属性，判断是否整个都在视口中
* `input()`方法增加`by_js`参数
* `save()`的`rename`参数改为`name`
* `get_src()`支持 blob 类型
* `css_path`获取的路径更精确
* 相对定位的`timeout`参数默认改为`None`
* `wait.delete()`方法改为`wait.deleted()`
* `wait.disabled_or_delete()`方法改为`wait.disabled_or_deleted()`
* `wait.display()`方法改为`wait.displayed()`
* 可用`==`比较两个元素
* 查找元素速度提高

---

✅️ 启动配置[​](#️-启动配置 "✅️ 启动配置的直接链接")
----------------------------------

### 📌 删除 easy\_set 方法[​](#-删除-easy_set-方法 "📌 删除 easy_set 方法的直接链接")

easy\_set 原本设计目的是为了方便修改 ini 文件设置。

本想着设置一次，之后无需再调用。但可能由于文档描述不清晰，很多人将其写到了正式的代码中。

因为 ini 文件的修改会影响其它项目，这是作者不  推荐的用法。

而且，实际使用中发现 easy\_set 并没有变得更方便，功能可用`ChromiumOptions`对象的`save()`方法代替。

因此决定将其删除，也可以免去多维护一份代码。

---

### 📌 支持设置实验项[​](#-支持设置实验项 "📌 支持设置实验项的直接链接")

4.0 新增支持启动浏览器时设置实验项，即`'chrome://flags'`中的项目。

使用新增的`set_flag()`方法设置。使用`clear_flags_in_file()`清空配置文件中已设置的项。

有哪些实验项，具体在`'chrome://flags'`查看。

**示例：**

```
from DrissionPage import ChromiumOptions  
  
co = ChromiumOptions()  
co.set_flag('temporary-unexpire-flags-m118', '1')  
co.set_flag('disable-accelerated-2d-canvas')
```

---

### 📌 ini 文件变化[​](#-ini-文件变化 "📌 ini 文件变化的直接链接")

* `chrome_options`类改为`chromium_options`
* `binary_location`项改为`browser_path`
* `page_load_strategy`项改为`load_mode`
* `debugger_address`项改为`address`
* `arguments`项删除`'--remote-allow-origins=*'`参数
* `arguments`项增加`'--no-default-browser-check'`、`'--disable-suggestions-ui'`、`'--disable-popup-blocking'`、`'--hide-crash-restore-bubble'`、`'--disable-features=PrivacySandboxSettings4'`参数
* `paths`类增加`tmp_paht`项
* 删除`experimental_options`项
* `chrome_options`类增加`prefs`、`flags`、`existing_only`项
* 增加`others`类，包含`retry_times`和`retry_interval`项

---

### 📌 `ChromiumOptions` 修改[​](#-chromiumoptions-修改 "-chromiumoptions-修改的直接链接")

* 增加`set_flag()`和`clear_flags_in_file()`，用于设置实验项
* 增加`existing_only()`方法和`is_existing_only`属性，可指定只接管浏览器而不自动启动新的
* 增加`ignore_certificate_errors()`方法，可忽略证书报错
* 增加`retry_times`、`retry_interval`属性和`set_retry()`方法，可设置重试参数
* 增加`incognito()`方法，可设置无痕模式
* 增加`set_tmp_path()`方法，可指定临时文件夹路径
* 增加`tmp_path`和`is_auto_port`属性
* `auto_port()`增加`tmp_path`参数
* `set_paths()`
  方法拆分成`set_browser_path()`、`set_local_port()`、`set_address()`、`set_download_path()`、`set_user_data_path()`、`set_cache_path()`
  方法
* `set_page_load_strategy()`改成`set_load_mode()`
* `set_headless()`改成`headless()`
* `set_no_imgs()`改成`no_imgs()`
* `set_no_js()`改成`no_js()`
* `set_mute()`改成`mute()`
* `debugger_address`改成`address`

### 📌 `SessionOptions` 修改[​](#-sessionoptions-修改 "-sessionoptions-修改的直接链接")

* `SessionOptions`的`set_paths()`方法改为`set_download_path()`
* 增加`retry_times`、`retry_interval`属性和`set_retry()`方法，可设置重试参数

---

### 📌 其它[​](#-其它-1 "📌 其它的直接链接")

* 启动或接管浏览器时，可自动关闭弹出的隐私声明
* 在无界面系统启动浏览器时，自动使用无头，可用`set_headless(False)`禁用
* 当`set_headless(False)`但接管了无头浏览器，将关闭并启动新的有头浏览器
* `auto_port()`方法支持多线程

---

✅️ 其它[​](#️-其它 "✅️ 其它的直接链接")
----------------------------

### 📌 删除 2.x 代码[​](#-删除-2x-代码 "📌 删除 2.x 代码的直接链接")

`MixPage` 是 `WebPage`的前身，基于 selenium。

随着 3.x 版本迭代，自研的底层日渐成熟并完全超越旧版。旧版已到了退休的时候。。

为对项目进行精简，避免新代码受旧功能制约，因此将旧代码删除。

删除以下类：`MixPage`、`DriverPage`、`DriverOptions`、`Drission`。

旧版对作者的成长有过巨大贡献，因此将其独立成一个库，以此来继续它的生命，和纪念它做出过的成果。

可用以下命令安装旧版尝旧：

```
pip install MixPage
```

---

### 📌 异常变化[​](#-异常变化 "📌 异常变化的直接链接")

* `CallMethodError`改为`CDPError`
* `ElementLossError`改为`ElementLostError`
* `ContextLossError`改为`ContextLostError`
* `TabClosedError`改为`PageDisconnectedError`
* 增加`WaitTimeoutError`
* 增加`GetDocumentError`
* 增加`WrongURLError`
* 增加`StorageError`
* 增加`CookieFormatError`
* 增加`TargetNotFoundError`

---

### 📌 Settings 变化[​](#-settings-变化 "📌 Settings 变化的直接链接")

* 增加`singleton_tab_obj`，设置 Tab 对象是否允许多例
* `raise_ele_not_found`改为`raise_when_ele_not_found`
* `raise_click_failed`改为`raise_when_click_failed`

---

### 📌 更多[​](#-更多-1 "📌 更多的直接链接")

* `handle_alert()`方法增加`next_one`参数，可处理下一个出现的弹窗
* 浏览器页面对象增加`set.auto_handle_alert()`方法，可设置自动处理弹窗
* `SessionPage`增加`set.encoding()`方法和`encoding`属性
* `<option>` 元素可以接受点击，操作更符合直觉
* `run_js()`、`run_js_loaded()`、`run_async_js()`方法增加`timeout`参数
* `run_async_js()`删除`timeout`参数
* `timeouts`的`implicit`改成`base`
* `ActionChains`改成`Actions`
* 动作链的移动方法增加`duration`参数
* 动作链增加`input()`方法
* 动作链`key_down()`和`key_up()`方法可接收按键名称文本
* 动作链`type()`方法`text`参数改为`keys`
* `get_screenshot()`方法增加`name`属性，可指定文件名
* 元素的`get_screenshot()`方法增加`scroll_to_center`参数，截图前先滚动到页面正中
* `wait.new_tab()`方法成功时返回新标签页 id
* `tabs`不包含 F12 的窗口
* `DrissionPage.common`路径增加`wait_until()`方法，支持自定义组合等待条件
* `DrissionPage.common`路径增加`get_blob()`方法，用于获取指定 blob 内容

---

✅️ 问题修复[​](#️-问题修复 "✅️ 问题修复的直接链接")
----------------------------------

* 修复网络连接极不稳定时获取文档失败问题
* 修复相对定位`timeout`失效问题
* 修复 shadow root 内定位元素可能偏差问题
* 修复异域`ChromiumFrame`内部元素无法获取屏幕坐标的问题
* 修复相对路径插件加载失败的问题
* 所有循环增加超时设置，避免出现卡死
* 修复元素截图时窗口外部分空白问题
* 修复 Tab 没有继承 Page 下载路径的问题
* 修复 `<iframe>` 内元素获取 href 属性错误问题
* 修复 cookie 设置 expires 时的问题

[上一页

💥 4.1 功能介绍](/features/4.1)[下一页

💥 3.2 功能介绍](/features/3)

* [✅️ 新的抓包功能](#️-新的抓包功能)
  + [📌 旧 api 变化](#-旧-api-变化)
  + [📌 新 api](#-新-api)
  + [📌 示例](#-示例)
* [✅️ 新的页面访问逻辑](#️-新的页面访问逻辑)
  + [📌 api 变化](#-api-变化)
  + [📌 行为变化](#-行为变化)
  + [📌 新的`none`加载模式](#-新的none加载模式)
  + [📌 示例](#-示例-1)
* [✅️ 新的下载管理功能](#️-新的下载管理功能)
  + [📌 api 变化](#-api-变化-1)
  + [📌 新增功能](#-新增功能)
  + [📌 行为变化](#-行为变化-1)
  + [📌 示例](#-示例-2)
* [✅️ 页面对象](#️-页面对象)
  + [📌 启动参数变化](#-启动参数变化)
  + [📌 内置动作链](#-内置动作链)
  + [📌 状态信息](#-状态信息)
  + [📌 其它](#-其它)
* [✅️ cookies 设置](#️-cookies-设置)
* [✅️ 标签页管理](#️-标签页管理)
  + [📌 不再支持`to_tab()`功能](#-不再支持to_tab功能)
  + [📌 `new_tab()`的新功能](#-new_tab的新功能)
  + [📌 标签页的位置与大小](#-标签页的位置与大小)
  + [📌 标签页的行为](#-标签页的行为)
  + [📌 `get_tab()`参数变化](#-get_tab参数变化)
* [✅️ 元素相关](#️-元素相关)
  + [📌 定位语法变化](#-定位语法变化)
  + [📌 相对定位参数优化](#-相对定位参数优化)
  + [📌 位置和大小](#-位置和大小)
  + [📌 点击](#-点击)
  + [📌 查找元素失败显示细节](#-查找元素失败显示细节)
  + [📌 设置查找失败元素返回默认值](#-设置查找失败 元素返回默认值)
  + [📌 更多](#-更多)
* [✅️ 启动配置](#️-启动配置)
  + [📌 删除 easy\_set 方法](#-删除-easy_set-方法)
  + [📌 支持设置实验项](#-支持设置实验项)
  + [📌 ini 文件变化](#-ini-文件变化)
  + [📌 `ChromiumOptions` 修改](#-chromiumoptions-修改)
  + [📌 `SessionOptions` 修改](#-sessionoptions-修改)
  + [📌 其它](#-其它-1)
* [✅️ 其它](#️-其它)
  + [📌 删除 2.x 代码](#-删除-2x-代码)
  + [📌 异常变化](#-异常变化)
  + [📌 Settings 变化](#-settings-变化)
  + [📌 更多](#-更多-1)
* [✅️ 问题修复](#️-问题修复)

* 🌟 特性演示
* ⭐ 模式切换

⭐ 模式切换
======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

用浏览器登录网站，然后切换到 requests 读取网页。两者会共享登录信息。

```
from DrissionPage import Chromium  
  
# 创建页面对象  
tab = Chromium().latest_tab    
# 访问个人中心页面（未登录，重定向到登录页面）  
tab.get('https://gitee.com/profile')    
  
# 输入账号密码登录  
tab.ele('@id:user_login').input('您的用户名')    
tab.ele('@id:user_password').input('您的密码\n')  
tab.wait.load_start()  
  
# 切换到 s 模式  
tab.change_mode()    
# 登录后 session 模式的输出  
print('登录后title：', tab.title, '\n')
```

**输出：**

```
登录后title： 个人资料 - 码云 Gitee.com
```

[上一页

⭐ 与 selenium 对比](/features/features_demos/selenium)[下一页

⭐ 获取元素属性](/features/features_demos/get_ele_attr)

* 🌟 特性演示
* ⭐ 获取元素属性

⭐ 获取元素属性
========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

```
foot = tab.ele('#footer-left')  # 用 id 查找元素  
first_col = foot.ele('css:>div')  # 使用 css selector 在元素的下级中查找元素（第一个）  
lnk = first_col.ele('text:命令学')  # 使用文本内容查找元素  
text = lnk.text  # 获取元素文本  
href = lnk.attr('href')  # 获取元素属性值  
  
print(text, href, '\n')  
  
# 简洁模式串联查找  
text = tab('@id:footer-left')('css:>div')('text:命令学').text  
print(text)
```

**输出：**

```
Git 命令学习 https://oschina.gitee.io/learn-git-branching/  
  
Git 命令学习
```

[上一页

⭐ 模式切换](/features/features_demos/change_mode)[下一页

⭐ 下载文件](/features/features_demos/download)

* 🌟 特性演示
* ⭐ 与 requests 对比

⭐ 与 requests 对比
===============

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

以下代码实现一模一样的功能，对比两者的代码量：

🔸 获取元素内容

```
url = 'https://baike.baidu.com/item/python'  
  
# 使用 requests：  
import requests  
from lxml import etree  
headers = {'User-Agent':'Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2272.118 Safari/537.36'}  
response = requests.get(url, headers = headers)  
html = etree.HTML(response.text)  
element = html.xpath('//h1')[0]  
title = element.text  
  
# 使用 DrissionPage：  
from DrissionPage import SessionPage  
page = SessionPage()  
page.get(url)  
title = page('tag:h1').text
```

Tips

DrissionPage 自带默认 headers

🔸 下载文件

```
url = 'https://www.baidu.com/img/flexible/logo/pc/result.png'  
save_path = r'C:\download'  
  
# 使用 requests：  
import requests  
r = requests.get(url)  
with open(f'{save_path}\\img.png', 'wb') as fd:  
   for chunk in r.iter_content():  
       fd.write(chunk)  
  
# 使用 DrissionPage：  
from DrissionPage import SessionPage  
page = SessionPage()  
page.download(url, save_path, 'img')  # 支持重命名，处理文件名冲突
```

[上一页

💖 特性](/features/)[下一页

⭐ 与 selenium 对比](/features/features_demos/selenium)

* 🌟 特性演示
* ⭐ 与 selenium 对比

⭐ 与 selenium 对比
===============

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

以下代码实现一模一样的功能，对比两者的代码量：

🔸 用显性等待方式查找第一个文本包含 some text 的元素。

```
# 使用 selenium：  
element = WebDriverWait(driver).until(ec.presence_of_element_located((By.XPATH, '//*[contains(text(), "some text")]')))  
  
# 使用 DrissionPage：  
element = tab('some text')
```

🔸 跳转到一个标签页

```
# 使用 selenium：  
driver.switch_to.window(driver.window_handles[0])  
  
# 使用 DrissionPage：  
tab = browser.latest_tab
```

🔸 按文本选择下拉列表

```
# 使用 selenium：  
from selenium.webdriver.support.select import Select  
select_element = Select(element)  
select_element.select_by_visible_text('text')  
  
# 使用 DrissionPage：  
element.select('text')
```

🔸 拖拽一个元素

```
# 使用 selenium：  
ActionChains(driver).drag_and_drop(ele1, ele2).perform()  
  
# 使用 DrissionPage：  
ele1.drag_to(ele2)
```

🔸 滚动窗口到底部（保持水平滚动条不变）

```
# 使用 selenium：  
driver.execute_script("window.scrollTo(document.documentElement.scrollLeft, document.body.scrollHeight);")  
  
# 使用 DrissionPage：  
tab.scroll.to_bottom()
```

🔸 设置 headless 模式

```
# 使用 selenium：  
options = webdriver.ChromeOptions()  
options.add_argument("--headless")  
  
# 使用 DrissionPage：  
options = ChromiumOptions().headless()
```

🔸 获取伪元素内容

```
# 使用 selenium：  
text = webdriver.execute_script('return window.getComputedStyle(arguments[0], "::after").getPropertyValue("content");', element)  
  
# 使用 DrissionPage：  
text = element.pseudo.after
```

🔸 获取 shadow-root

新版 selenium 已可直接获取 shadow-root，但生成的 ShadowRoot 对象功能实在是太少了。

```
# 使用 selenium：  
shadow_element = webdriver.execute_script('return arguments[0].shadowRoot', element)  
  
# 使用 DrissionPage：  
shadow_element = element.sr
```

🔸 用 xpath 直接获取属性或文本节点（返回文本）

```
# 使用 selenium：  
相当复杂  
  
# 使用 DrissionPage：  
class_name = element('xpath://div[@id="div_id"]/@class')  
text = element('xpath://div[@id="div_id"]/text()[2]')
```

[上一页

⭐ 与 requests 对比](/features/features_demos/requests)[下一页

⭐ 模式切换](/features/features_demos/change_mode)

## <a name="入门"></a>入门

* 🌏 准备工作

本页总览

🌏 准备工作
======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

开始前，我们先设置一下浏览器路径。

如果只使用收发数据包功能，无需任何准备工作。

程序默认控制 Chrome，所以下面用 Chrome 演示。
如果要使用 Edge 或其它 Chromium 内核浏览器，设置方法是一样的。

注意

尽量使用版本号在 100 以上的浏览器，旧版有些功能不支持。

1️⃣ 尝试启动浏览器[​](#1️⃣-尝试启动浏览器 "1️⃣ 尝试启动浏览器的直接链接")
-----------------------------------------------

默认状态下，程序会自动在系统内查找 Chrome 路径。

执行以下代码，浏览器启动并且访问了项目官网，说明可直接使用，跳过后面的步骤即可。

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.get('http://DrissionPage.cn')
```

---

2️⃣ 设置路径[​](#2️⃣-设置路径 "2️⃣ 设置路径的直接链接")
--------------------------------------

如果上面的步骤提示出错，说明程序没在系统里找到 Chrome 浏览器。

可用以下其中一种方法设置，设置会持久化记录到默认配置文件，之后程序会使用该设置启动。

获取浏览器路径的方法

* 这里的浏览器路径不一定是 Chrome，Edge 等 Chromium 内核的浏览器都可以。
* 打开浏览器，在地址栏输入`chrome://version`（Edge 输入`edge://version`），回车。
  ![](/assets/images/find_browser_path-1b46b5e4ba053115091a598d3b4211ac.png)  
  如图所示，红框中就是要获取的路径。  
  此法不限于 Windows，有界面的 Linux 也可使用。

📌 方法一[​](#-方法一 "📌 方法一的直接链接")
----------------------------

新建一个临时 py 文件，并输入以下代码，填  入您电脑里的 Chrome 浏览器可执行文件路径，然后运行。

```
from DrissionPage import ChromiumOptions  
  
path = r'D:\Chrome\Chrome.exe'  # 请改为你电脑内Chrome可执行文件路径  
ChromiumOptions().set_browser_path(path).save()
```

这段代码会把浏览器路径记录到配置文件，今后启动浏览器皆使用该路径。

如果是想临时切换浏览器路径以尝试运行和操作是否正常，可以去掉`.save()`，以如下方式结合第1️⃣步的代码。

```
from DrissionPage import Chromium, ChromiumOptions  
  
path = r'D:\Chrome\Chrome.exe'  # 请改为你电脑内Chrome可执行文件路径  
co = ChromiumOptions().set_browser_path(path)  
tab = Chromium(co).latest_tab  
tab.get('http://DrissionPage.cn')
```

📌 方法二[​](#-方法二 "📌 方法二的直接链接")
----------------------------

在命令行输入以下命令（路径改成自己电脑里的）：

```
dp -p "D:\Chrome\chrome.exe"
```

注意

* 注意命令行的 python 环境与项目应是同一个
* 注意要先使用 cd 命令定位到项目路径

---

3️⃣ 重试控制浏览器[​](#3️⃣-重试控制浏览器 "3️⃣ 重试控制浏览器的直接链接")
-----------------------------------------------

现在，请重新执行第1️⃣步的代码，如果正确访问了项目官网，说明已经设置完成。

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
tab.get('http://DrissionPage.cn')
```

---

✅️️ 说明[​](#️️-说明 "✅️️ 说明的直接链接")
-------------------------------

当您完成准备工作后，无需关闭浏览器，后面的上手示例可继续接管当前浏览器。

[上一页

🌏 设置语言 / Set Language](/get_start/set_lang)[下一页

🗺️ 自动登录](/get_start/examples/control_browser)

* [1️⃣ 尝试启动浏览器](#1️⃣-尝试启动浏览器)
* [2️⃣ 设置路径](#2️⃣-设置路径)
* [📌 方法一](#-方法一)
* [📌 方法二](#-方法二)
* [3️⃣ 重试控制浏览器](#3️⃣-重试控制浏览器)
* [✅️️ 说明](#️️-说明)

* ☀ ️ 基本概念

本页总览

☀️ 基本概念
=======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

本节讲解 DrissionPage 的一些基本概念。了解它大概的构成。

如果您觉得有点懵，可直接跳过本节。

✅️️ 网页自动化[​](#️️-网页自动化 "✅️️ 网页自动化的直接链接")
----------------------------------------

网页自动化的形式通常有两种，它们各有优劣：

* 直接向服务器发送数据包，获取需要的数据
* 控制浏览器跟网页进行交互

前者轻量级，速度快，便于多线程、分布式部署，如 requests 库。但当数据包构成复杂，甚至加入加密技术时，开发过程烧脑程度直线上升。

鉴于此，DrissionPage 以页面为单位将两者整合，对 Chromium 协议 和 requests 进行了重新封装，实现两种模式的互通，并加入常用的页面和元素控制功能，可大幅降低开发难度和代码量。  
用于操作浏览器的对象叫 Driver，requests 用于管理连接的对象叫 Session，Drission 就是它们两者的合体。Page 表示以页面为单位使用。

在旧版本，本库是通过对 selenium 和 requests 的重新封装实现的。  
从 3.0 版开始，作者另起炉灶，自行实现了 selenium 全部功能，从而摆脱了对 selenium 的依赖，功能更多更强，运行效率更高，开发更灵活。
4.0 则在 3.0 经验的基础上对整个项目底层进行了重构，逻辑更合理。

如果您想了解旧版，请查 阅“旧版使用方法”章节。

---

✅️️ 基本使用逻辑[​](#️️-基本使用逻辑 "✅️️ 基本使用逻辑的直接链接")
-------------------------------------------

无论是控制浏览器，还是收发数据包，其操作逻辑是一致的。

即先创建页面对象，然后从页面对象中获取元素对象，通过对元素对象的读取或操作，实现数据的获取或页面的控制。

因此，最主要的对象就是两种：页面对象，及其生成的元素对象。

---

✅️️ 主要对象[​](#️️-主要对象 "✅️️ 主要对象的直接链接")
-------------------------------------

### 📌 浏览器和标签页对象[​](#-浏览器和标签页对象 "📌 浏览器和标签页对象的直接链接")

* `Chromium`：浏览器对象，用于连接浏览器、管理标签页及其它和浏览器总体有关的操作
* `MixTab`：浏览器标签页对象，由`Chromium`对象产生，一个对象控制一个实际的标签页
* `ChromiumTab`：和`MixTab`一样也是标签页对象，由`ChromiumPage`对象产生，不可切换收发数据包模式

### 📌 元素对象[​](#-元素对象 "📌 元素对象的直接链接")

* `ChromiumElement`：浏览器元素对象
* `SessionElement`：静态元素对象
* `ChromiumFrame`：`<iframe>`元素对象，兼有标签页对象和元素特性
* `ShadowRoot`：shadow-root 元素对象

### 📌 Page 对象[​](#-page-对象 "📌 Page 对象的直接链接")

* `ChromiumPage`：能管理浏览器本身的标签页对象，可用作程序入口
* `WebPage`：类似于`ChromiumPage`，整合浏览器控制和收发数据包于一体的页面对象
* `SessionPage`：单纯用于收发数据包的页面对象，可单独使用

### 📌 称呼[​](#-称呼 "📌 称呼的直接链接")

文档里经常用到这几个称呼：

* `MixTab`、`ChromiumTab`统称为 Tab 对象
* `ChromiumPage`、`WebPage`和`SessionPage`统称为 Page 对象
* Page 对象、Tab 对象和`ChromiumFrame`统称为页面对象

---

  

✅️️ 对象关系图[​](#️️-对象关系图 "✅️️ 对象关系图的直接链接")
----------------------------------------

下图列出本库中要用到的各种对象的生成关系。

```
├─ SessionPage  
|     └─ SessionElement  
|           └─ SessionElement  
├─ Chrmoium  
|     └─ MixTab  
|           ├─ ChromiumElement  
|           |    ├─ ChromiumElement  
|           |    ├─ ChromiumFrame  
|           |    └─ SessionElement  
|           ├─ SessionElement  
|           |    └─ SessionElement  
|           ├─ ChromiumFrame  
|           |    ├─ ChromiumElement  
|           |    ├─ ChromiumFrame  
|           |    └─ SessionElement  
|           └─ ShadowRoot  
|                ├─ ChromiumElement  
|                ├─ ChromiumFrame  
|                └─ SessionElement  
├─ SessionOptions  
└─ ChrmoiumOptions
```

---

✅️️ 工作模式[​](#️️-工作模式 "✅️️ 工作模式的直接链接")
-------------------------------------

`MixTab`和`WebPage`既可控制浏览器，也可用数据包方式访问网络数据。
它们有两种工作方式：d 模式和 s 模式。  
页面对象可以在这两种模式间切换，两种模式拥有一致的使用方法，但任一时间只能处于其中一种模式。

### 📌 d 模式[​](#-d-模式 "📌 d 模式的直接链接")

d 模式既表示 Driver，还有 Dynamic 的意思。  
d 模式用于控制浏览器，不仅可以读取浏览器获取到的信息，还能对页面进  行操作，如点击、填写、开关标签页、改变元素属性、执行 js 脚本等等。  
d 模式功能强大，但运行速度受浏览器制约非常缓慢，而且需要占用大量内存。

---

### 📌 s 模式[​](#-s-模式 "📌 s 模式的直接链接")

s 模式既表示 Session，还有 speed、silence 的意思。  
s 模式的运行速度比 d 模式快几个数量级，但只能基于数据包进行读取或发送，不能对页面进行操作，不能运行 js。  
爬取数据时，如网站数据包较为简单，应首选 s 模式。

---

### 📌 模式切换[​](#-模式切换 "📌 模式切换的直接链接")

`MixTab`和`WebPage`对象可以在 d 模式和 s 模式之间切换，这通常用于以下情况：

* 当登录验证很严格，难以解构，如有验证码的时候，用浏览器处理登录，然后转换成 s 模式爬取数据。既避免了处理烧脑的 js，又能享受 s 模式的速度。
* 页面数据由 js 产生，且页面结构极其复杂，可以用 d 模式读取页面元素，然后把元素转成 s 模式的元素进行分析。可以极大地提高 d 模式的处理速度。

---

✅️️ 配置管理[​](#️️-配置管理 "✅️️ 配置管理的直接链接")
-------------------------------------

无论 requests 还是浏览器，都通常需要一些配置信息才能正常工作，如长长的`user_agent`、浏览器 exe 文件路径、浏览器配置等。
这些代码往往是繁琐而重复的，不利于代码  的简洁。  
因此，DrissionPage 使用配置文件记录常用配置信息，程序会自动读取默认配置文件里的内容。
所以，在示例中，通常看不见配置信息的代码。

这个功能支持用户保存不同的配置文件，按情况调研，也可以支持直接把配置写在代码里面，屏蔽读取配置文件。

Tips

当需要打包程序时，必需把配置写到代码里，或打包后手动复制配置文件到运行路径，否则会报错。详见相关章节。

### 📌 `SessionOptions`[​](#-sessionoptions "-sessionoptions的直接链接")

用于`SessionPage`和`WebPage` s 模式的配置对象。

---

### 📌 `ChromiumOptions`[​](#-chromiumoptions "-chromiumoptions的直接链接")

用于用于浏览器的配置对象。

---

✅️️ 定位符[​](#️️-定位符 "✅️️ 定位符的直接链接")
----------------------------------

定位符用于定位页面中的元素，是本库一大特色，能够用非常简明的方式来获取元素，简洁易用。
可读性和易用性高于 xpath 等其它方式，并且兼容 xpath、css selector、selenium 定位符。

以下是一组对比：

定位文本包含`'abc'`的元素：

```
# DrissionPage  
ele = tab('abc')  
  
# selenium  
ele = driver.find_element(By.XPATH, '//*[contains(text(), "abc"]')
```

定位 class 为`'abc'`的元素：

```
# DrissionPage  
ele = tab('.abc')  
  
# selenium  
ele = driver.find_element(By.CLASS_NAME, 'abc')
```

定位 ele 元素的兄弟元素：

```
# DrissionPage  
ele1 = ele.next()  # 获取后一个元素  
ele1 = ele.prev(index=2)  # 获取前面第二个元素  
  
# selenium  
ele1 = ele.find_element(By.XPATH, './/following-sibling::*')  # 获取有i一个元素  
ele1 = ele.find_element(By.XPATH, './/preceding-sibling::*[2]')  # 获取前面第二个元素
```

显然，本库的定位语句更简洁易懂，还有很多灵活好用的方法，详见 “查找元素” 章节。

[上一页

🗺️ 模式切换](/get_start/examples/switch_mode)

* [✅️️ 网页自动化](#️️-网页自动化)
* [✅️️ 基本使用逻辑](#️️-基本使用逻辑)
* [✅️️ 主要对象](#️️-主要对象)
  + [📌 浏览器和标签页对象](#-浏览器和标签页对象)
  + [📌 元素对象](#-元素对象)
  + [📌 Page 对象](#-page-对象)
  + [📌 称呼](#-称呼)
* [✅️️ 对象关系图](#️️-对象关系图)
* [✅️️ 工作模式](#️️-工作模式)
  + [📌 d 模式](#-d-模式)
  + [📌 s 模式](#-s-模式)
  + [📌 模式切换](#-模式切换)
* [✅️️ 配置管理](#️️-配置管理)
  + [📌 `SessionOptions`](#-sessionoptions)
  + [📌 `ChromiumOptions`](#-chromiumoptions)
* [✅️️ 定位符](#️️-定位符)

* 🌏 上手示例
* 🗺️ 自动登录

本页总览

🗺️ 自动登录
=======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

本示例演示控制浏览器登录 gitee 网站。

✅️️ 页面分析[​](#️️-页面分析 "✅️️ 页面分析的直接链接")
-------------------------------------

网址：<https://gitee.com/login>

打开网址，按`F12`，我们可以看到页面 html 如下：

![](/assets/images/gitee_1-a0c506a07b80ba8c2713bf306a475c2e.jpg)

用户名输入框`id`为`'user_login'`，密码输入框`id`为`'user_password'`，登录按钮`value`为`'登 录'`。

我们可以用这三个属性定位这三个元素，然后对其输入数据和点击。

---

✅️️ 示例代码[​](#️️-示例代码 "✅️️ 示例代码的直接链接")
-------------------------------------

您可以把以下代码复制到编辑器，把账号和密码改为您自己的，可直接执行看到运行结果。

```
from DrissionPage import Chromium  
  
# 启动或接管浏览器，并获取标签页对象  
tab = Chromium().latest_tab  
# 跳转到登录页面  
tab.get('https://gitee.com/login')  
  
# 定位到账号文本框，获取文本框元素  
ele = tab.ele('#user_login')  
# 输入对文本框输入账号  
ele.input('您的账号')  
# 定位到密码文本框并输入密码  
tab.ele('#user_password').input('您的密码')  
# 点击登录按钮  
tab.ele('@value=登 录').click()
```

---

✅️️ 示例详解[​](#️️-示例详解 "✅️️ 示例详解的直接链接")
-------------------------------------

我们逐行解读代码：

```
from DrissionPage import Chromium
```

↑ 首先，我们导入用于控制浏览器的类`Chromium`。

```
tab = Chromium().latest_tab
```

↑ 接下来，我们创建一个`Chromium`对象，用于连接浏览器，并用`latest_tab`获取一个标签页对象。

```
tab.get('https://gitee.com/login')
```

↑ `get()`方法用于访问参数中的网址。它会等待页面完全加载，再继续执行后面的代码。
您也可以修改等待策略，如等待 DOM 加载而不等待资源下载，就停止加载，这将在后面的章节说明。

```
ele = tab.ele('#user_login')
```

↑ `ele()`方法用于查找元素，它返回一个`ChromiumElement`对象，用于操作元素。

`'#user_login'`是定位符文本，`#`意思是按`id`属性查找元素。

值得一提的是，`ele()`内置了等待，如果元素未加载，它会执行等待，直到元素出现或到达时限。默认超时时间 10 秒。

```
ele.input('您的账号')
```

↑ `input()`方法用于对元素输入文本。

```
tab.ele('#user_password').input('您的密码')
```

↑ 我们也可以进行链式操作，获取元素后直接输入文本。

```
tab.ele('@value=登 录').click()
```

↑ 输入账号密码后，以相同的方法获取按钮元素，并对其执行点击操作。

不同的是，这次通过其`value`属性作为查找条件。`@`表示按属性名查找。

到这里，我们已完成了自动登录 gitee 网站的操作。

[上一页

🌏 准备工作](/get_start/before_start)[下一页

🗺️ 收发数据包](/get_start/examples/data_packets)

* [✅️️ 页面分析](#️️-页面分析)
* [✅️️ 示例代码](#️️-示例代码)
* [✅️️ 示例详解](#️️-示例详解)

* 🌏 上手示例
* 🗺️ 收发数据包

本页总览

🗺️ 收发数据包
========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

本示例演示用`SessionPage`已收发数据包的方式采集 gitee 网站数据。

本示例不使用浏览器。

✅️️ 页面分析[​](#️️-页面分析 "✅️️ 页面分析的直接链接")
-------------------------------------

网址：<https://gitee.com/explore/all>

这个示例的目标，要获取所有库的名称和链接，为避免对网站造成压力，我们只采集 3 页。

打开网址，按`F12`，我们可以看到页面 html 如下：

![](/assets/images/gitee_2-a46709588060c6c79e04dd943765307a.jpg)

从 html 代码中可以看到，所有开源项目的标题都是`class`属性为`'title project-namespace-path'`的`<a>`元素。我们可以遍历这些`<a>`元素，获取它们的信息。

同时，我们观察到，列表页网址是以页数为参数访问的，如第一页 url 为`https://gitee.com/explore/all?page=1`，页数就是`page`参数。因此我们可以通过修改这个参数访问不同的页面。

---

✅️️ 示例代码[​](#️️-示例代码 "✅️️ 示例代码的直接链接")
-------------------------------------

以下代码可直接运行并查看结果：

```
from DrissionPage import SessionPage  
  
# 创建页面对象  
page = SessionPage()  
  
# 爬取3页  
for i in range(1, 4):  
    # 访问某一页的网页  
    page.get(f'https://gitee.com/explore/all?page={i}')  
    # 获取所有开源库<a>元素列表  
    links = page.eles('.title project-namespace-path')  
    # 遍历所有<a>元素  
    for link in links:  
        # 打印链接信息  
        print(link.text, link.link)
```

**输出：**

```
小熊派开源社区/BearPi-HM_Nano https://gitee.com/bearpi/bearpi-hm_nano  
明月心/PaddleSegSharp https://gitee.com/raoyutian/PaddleSegSharp  
RockChin/QChatGPT https://gitee.com/RockChin/QChatGPT  
TopIAM/eiam https://gitee.com/topiam/eiam  
  
以下省略。。。
```

---

✅️️ 示例详解[​](#️️-示例详解 "✅️️ 示例详解的直接链接")
-------------------------------------

我们逐行解读代码：

```
from DrissionPage import SessionPage
```

↑ 首先，我们导入用于收发数据包的页面类`SessionPage`。

```
page = SessionPage()
```

↑ 接下来，我们创建一  个`SessionPage`对象。

```
for i in ranage(1, 4):  
    page.get(f'https://gitee.com/explore/all?page={i}')
```

↑ 然后我们循环 3 次，以构造每页的 url，每次都用`get()`方法访问该页网址。

```
    links = page.eles('.title project-namespace-path')
```

↑ 访问网址后，我们用页面对象的`eles()`获取页面中所有`class`属性为`'title project-namespace-path'`的元素对象。

`eles()`方法用于查找多个符合条件的元素，返回由它们组成的`list`。

这里查找的条件是`class`属性，`.`表示按`class`属性查找元素。

```
    for link in links:  
        print(link.text, link.link)
```

↑ 最后，我们遍历获取到的元素列表，获取每个元素的属性并打印出来。

`.text`获取元素的文本，`.link`获取元素的`href`或`src`属性。

[上一页

🗺️ 自动登录](/get_start/examples/control_browser)[下一页

🗺️ 模式切换](/get_start/examples/switch_mode)

* [✅️️ 页面分析](#️️-页面分析)
* [✅️️ 示例代码](#️️-示例代码)
* [✅️️ 示例详解](#️️-示例详解)

* 🌏 上手示例
* 🗺️ 模式切换

本页总览

🗺️ 模式切换
=======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

这个示例演示标签页对象如何切换控制浏览器和收发数据包两种模式。

通常，切换模式是用来应付登录检查很严格的网站，可以用浏览器处理登录，再转换模式用收发数据包的形式来采集数据。
但是这种场景需要有对应的账号，不便于演示。演示使用浏览器在 gitee 搜索，然后转换到收发数据包的模式来读取数据。
虽然此示例现实使用意义不大，但可以了解其工作模式。

✅️️ 页面分析[​](#️️-页面分析 "✅️️ 页面分析的直接链接")
-------------------------------------

网址：<https://gitee.com/explore>

打开网址，按`F12`，我们可以看到页面 html 如下：

![](/assets/images/change1-bac86dfcdb35ce852aaa00390c5100f1.png)

输入框`<input>`元素`id`属性为`'q'`，搜索按钮`<button>`元素`文本`包含`'搜索'`文本，可用来作条件查找元素。

输入关键词搜索后，再查看页面 html：

![](/assets/images/change2-5177047ea61c4c245d44736eeb34d8fd.png)

通过分析 html 代码，我们可以看出，每个结果的标题都存在`id`为`'hits-list'`里面，`class`为`'item'`的元素中。因此，我们可以获取页面中所有这些元素，再遍历获取其信息。

---

✅️️ 示例代码[​](#️️-示例代码 "✅️️ 示例代码的直接链接")
-------------------------------------

您可以直接运行以下代码：

```
from DrissionPage import Chromium  
  
# 连接浏览器并获取一个MixTab对象  
tab = Chromium().latest_tab  
# 访问网址  
tab.get('https://gitee.com/explore/all')  
# 切换到收发数据包模式  
tab.change_mode()  
# 获取所有行元素  
items = tab.ele('.ui relaxed divided items explore-repo__list').eles('.item')  
# 遍历获取到的元素  
for item in items:  
    # 打印元素文本  
    print(item('t:h3').text)  
    print(item('.project-desc mb-1').text)  
    print()
```

**输出：**

```
dromara/Sa-Token  
一个轻量级 Java 权限认证框...  
  
lengleng/pig  
基于Spring Boot 3.3...  
  
...
```

---

✅️️ 示例详解[​](#️️-示例详解 "✅️️ 示例详解的直接链接")
-------------------------------------

我们逐行解读代码：

```
from DrissionPage import Chromium
```

↑ 首先，我们导入用于控制浏览器的类`Chromium`。

```
tab = Chromium().latest_tab
```

↑ 接下来，们创建一个`Chromium`对象，用于连接浏览器，并用`latest_tab`获取一个可切换模式的标签页对象。

```
tab.get('https://gitee.com/explore')
```

↑ 然后控制浏览器访问 gitee。

```
tab.change_mode()
```

↑ `change_mode()`方法用于切换工作模式，从当前控制浏览器的模式切换到收发数据包模式。

切换的时候程序会在新模式重 新访问当前 url。

```
items = tab('#hits-list').eles('.item')
```

↑ 切换后，我们可以用与控制浏览器一致的语法，获取页面元素，这获取页面中所有结果行素，它返回这些元素对象组成的列表。

```
for item in items:  
    print(item('.title').text)  
    print(item('.desc').text)  
    print()
```

↑ 最后，我们遍历这些元素，并逐个打印它们包含的文本。

[上一页

🗺️ 收发数据包](/get_start/examples/data_packets)[下一页

☀️ 基本概念](/get_start/concept)

* [✅️️ 页面分析](#️️-页面分析)
* [✅️️ 示例代码](#️️-示例代码)
* [✅️️ 示例详解](#️️-示例详解)

* 🌏 导入

本页总览

🌏 导入
====

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

DrissionPage 提供的功能放在以下几个路径：

* `from DrissionPage import ****`：浏览器类、配置类、页面类
* `from DrissionPage.errors import ****`：异常
* `from DrissionPage.common import ****`：辅助工具
* `from DrissionPage.items import ****`：衍生对象，用于类型判断

✅️ 浏览器类[​](#️-浏览器类 "✅️ 浏览器类的直接链接")
----------------------------------

### 📌 `Chromium`[​](#-chromium "-chromium的直接链接")

浏览器类用于连接浏览器、管理标签页及其它和浏览器总体有关的操作。

浏览器类相当于总管，它可以作为浏览器入口，使用它产生 Tab 对象去操控每个标签页。

```
from DrissionPage import Chromium
```

---

✅️ 页面类[​](#️-页面类 "✅️ 页面类的直接链接")
-------------------------------

### 📌 `ChromiumPage`[​](#-chromiumpage "-chromiumpage的直接链接")

`ChromiumPage`是将浏览器对象和第一个标签页对象封装在一起，用于控制浏览器。

`ChromiumPage`只是简化了操作，使用效果和直接使用`Chromium`对象基本一致。

唯一区别是，`ChromiumPage`生成的标签页对象是`ChromiumTab`，不能切换模式。

```
from DrissionPage import ChromiumPage
```

---

### 📌 `WebPage`[​](#-webpage "-webpage的直接链接")

`WebPage`与`ChromiumPage`相似，不过其自身及其产生的 Tab 对象可切换模式，既可控制浏览器，也可收发数据包。

```
from DrissionPage import WebPage
```

---

### 📌 `SessionPage`[​](#-sessionpage "-sessionpage的直接链接")

`SessionPage`用于收发数据包，是对 requests 和 lxml 进行封装实现的。

它把网络连接和结果解析封装成页面。操作逻辑和其它页面一致。

```
from DrissionPage import SessionPage
```

---

✅️ 配置工具[​](#️-配置工具 "✅️ 配置工具的直接链接")
----------------------------------

### 📌 `ChromiumOptions`[​](#-chromiumoptions "-chromiumoptions的直接链接")

`ChromiumOptions`类用于设置浏览器启动参数。

这些参数只有在启动浏览器时有用，接管已存在的浏览器时是不生效的。

```
from DrissionPage import ChromiumOptions
```

---

### 📌 `SessionOptions`[​](#-sessionoptions "-sessionoptions的直接链接")

`SessionOptions`类用于设置`Session`对象启动参数。

用于配置`SessionPage`或`WebPage`的 s 模式的连接参数。

```
from DrissionPage import SessionOptions
```

---

### 📌 `Settings`[​](#-settings "-settings的直接链接")

`Settings`用于设置全局运行配置，如找不到元素时是否抛出异常等。

```
from DrissionPage.common import Settings
```

---

  

✅️ 辅助工具[​](#️-辅助工具 "✅️ 辅助工具的直接链接")
----------------------------------

### 📌 `Keys`[​](#-keys "-keys的直接链接")

键盘按键类，用于键入 ctrl、alt 等按键。

```
from DrissionPage.common import Keys
```

---

### 📌 `By`[​](#-by "-by的直接链接")

与 selenium 一致的`By`类，便于项目迁移。

```
from DrissionPage.common import By
```

---

### 📌 其它工具[​](#-其它工具 "📌 其它工具的直接链接")

这些工具都在`DrissionPage.common`路径中。

* `wait_until`：可等待传入的方法结果为真
* `make_session_ele`：从 html 文本生成`ChromiumElement`对象
* `configs_to_here`：把配置文件复制到当前路径
* `get_blob`：获取指定的 blob 资源
* `tree`：用于打印页面对象或元素对象结构
* `from_selenium`：用于对接 selenium 代码
* `from_playwright`：用于对接 playwright 代码

```
from DrissionPage.common import wait_until  
from DrissionPage.common import make_session_ele  
from DrissionPage.common import configs_to_here
```

---

✅️ 异常[​](#️-异常 "✅️ 异常的直接链接")
----------------------------

异常放在`DrissionPage.errors`路径。

全部异常详见进阶使用章节。

```
from DrissionPage.errors import ElementNotFoundError
```

---

✅️ 衍生对象类型[​](#️-衍生对象类型 "✅️ 衍生对象类型的直接链接")
----------------------------------------

Tab、Element 等被其它对象生成的对象，开发过程中需要类型判断时需要导入这些类型。

可在`DrissionPage.items`路径导入。

```
from DrissionPage.items import SessionElement  
from DrissionPage.items import ChromiumElement  
from DrissionPage.items import ShadowRoot  
from DrissionPage.items import NoneElement  
from DrissionPage.items import ChromiumTab  
from DrissionPage.items import MixTab  
from DrissionPage.items import ChromiumFrame
```

[上一页

🌏 安装](/get_start/installation)[下一页

🌏 设置语言 / Set Language](/get_start/set_lang)

* [✅️ 浏览器类](#️-浏览器类)
  + [📌 `Chromium`](#-chromium)
* [✅️ 页面类](#️-页面类)
  + [📌 `ChromiumPage`](#-chromiumpage)
  + [📌 `WebPage`](#-webpage)
  + [📌 `SessionPage`](#-sessionpage)
* [✅️ 配置工具](#️-配置工具)
  + [📌 `ChromiumOptions`](#-chromiumoptions)
  + [📌 `SessionOptions`](#-sessionoptions)
  + [📌 `Settings`](#-settings)
* [✅️ 辅助工具](#️-辅助工具)
  + [📌 `Keys`](#-keys)
  + [📌 `By`](#-by)
  + [📌 其它工具](#-其它工具)
* [✅️ 异常](#️-异常)
* [✅️ 衍生对象类型](#️-衍生对象类型)

* 🌏 安装

本页总览

🌏 安装
====

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

✅️️ 运行环境[​](#️️-运行环境 "✅️️ 运行环境的直接链接")
-------------------------------------

操作系统：Windows、Linux 和 Mac

Python 版本：3.6 及以上

支持：Chromium 内核浏览器（如 Chrome 和 Edge）、electron 应用

---

✅️️ 安装[​](#️️-安装 "✅️️ 安装的直接链接")
-------------------------------

请使用 pip 安装：

```
pip install DrissionPage
```

---

✅️️ 升级[​](#️️-升级 "✅️️ 升级的直接链接")
-------------------------------

### 📌 升级最新稳定版[​](#-升级最新稳定版 "📌 升级最新稳定版的直接链接")

```
pip install DrissionPage --upgrade
```

---

### 📌 指定版本升级[​](#-指定版本升级 "📌 指定版本升级的直接链接")

```
pip install DrissionPage==4.0.0b17
```

[下一页

🌏 导入](/get_start/import)

* [✅️️ 运行环境](#️️-运行环境)
* [✅️️ 安装](#️️-安装)
* [✅️️ 升级](#️️-升级)
  + [📌 升级最新稳定版](#-升级最新稳定版)
  + [📌 指定版本升级](#-指定版本升级)

* 🌏 设置语言 / Set Language

本页总览

🌏 设置语言 / Set Language
=====================

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

DrissionPage 的报错信息及提示支持中文和英文。

DrissionPage error messages and prompts are available in Chinese and English.

✅️ 设置方法 / Usage[​](#️-设置方法--usage "✅️ 设置方法 / Usage的直接链接")
---------------------------------------------------------

```
from DrissionPage.common import Settings  
  
Settings.set_language('en')  # 设置为中文时，填入'zh_cn'
```

[上一页

🌏 导入](/get_start/import)[下一页

🌏 准备工作](/get_start/before_start)

* [✅️ 设置方法 / Usage](#️-设置方法--usage)

## <a name="SessionPage"></a>SessionPage

* 🛫 SessionPage
* 🛩️ 创建页面对象

本页总览

🛩️ 创建页面对象
=========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

✅️️ `SessionPage`初始化参数[​](#️️-sessionpage初始化参数 "️️-sessionpage初始化参数的直接链接")
--------------------------------------------------------------------------

| 初始化参数 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `session_or_options` | `Session` `SessionOptions` | `None` | 传入`Session`对象时使用该对象收发数据包；传入`SessionOptions`对象时用该配置创建`Session`对象；为`None`则从 ini 文件读取，为`False`则不从 ini 文件读取，而用内置默认配置 |

---

✅️️ 直接创建[​](#️️-直接创建 "✅️️ 直接创建的直接链接")
-------------------------------------

这种方式代码最简洁，程序会从配置文件中读取配置，自动生成页面对象。

```
from DrissionPage import SessionPage  
  
page = SessionPage()
```

`SessionPage`无需控制浏览器，无需做任何配置即可使用。

直接创建时，程序默认读取 ini 文件配置，如 ini 文件不存在，会使用内置配置。

默认 ini 和内置配置信息详见 “进阶使用->配置文件的使用” 章节。

---

  

✅️️ 通过配置信息创建[​](#️️-通过配置信息创建 "✅️️ 通过配置信息创建的直接链接")
-------------------------------------------------

如果需要在使用前进行一些配置，可使用`SessionOptions`。它是专门用于设置`Session`对象初始状态的类，内置了常用的配置。详细使用方法见 “启动配置” 一节。

### 📌 使用方法[​](#-使用方法 "📌 使用方法的直接链 接")

在`SessionPage`创建时，将已创建和设置好的`SessionOptions`对象以参数形式传递进去即可。

| 初始化参数 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `read_file` | `bool` | `True` | 是否从 ini 文件中读取配置信息 |
| `ini_path` | `str` | `None` | 文件路径，为`None`则读取默认 ini 文件 |

注意

`Session`对象创建后再修改这个配置是没有效果的。

```
# 导入 SessionOptions  
from DrissionPage import SessionPage, SessionOptions  
  
# 创建配置对象，并设置代理信息  
so = SessionOptions().set_proxies(http='127.0.0.1:1080')  
# 用该配置创建页面对象  
page = SessionPage(session_or_options=so)
```

Tips

您可以把配置保存到配置文件以后自动读取，详见 “启动配置” 章节。

---

### 📌 从指定 ini 文件创建[​](#-从指定-ini-文件创建 "📌 从指定 ini 文件创建的直接链接")

以上方法是使用默认 ini 文件中保存的配置信息创建对象，你可以保存一个 ini 文件到别的地方，并在创建对象时指定使用它。

```
from DrissionPage import SessionPage, SessionOptions  
  
# 创建配置对象时指定要读取的ini文件路径  
so = SessionOptions(ini_path=r'./config1.ini')  
# 使用该配置对象创建页面  
page = SessionPage(session_or_options=so)
```

---

### 📌 不使用 ini 文件[​](#-不使用-ini-文件 "📌 不使用 ini 文件的直接链接")

可以用以下方法，指定不使用 ini 文件的配置，而把配置写在代码中。

```
from DrissionPage import SessionPage, SessionOptions  
  
so = SessionOptions(read_file=False)  # read_file设为False  
so.set_retry(5)  
page = SessionPage(so)
```

---

✅️️ 传递控制权[​](#️️-传递控制权 "✅️️ 传递控制权的直接链接")
----------------------------------------

当需要使用多个页面对象共同操作一个页面时，可在页面对象创建时接收另一个页面间对象传递过来的`Session`对象，以达到多个页面对象同时使用一个`Session`对象的效果。

```
from DrissionPage import SessionPage  
  
# 创建一个页面  
page1 = SessionPage()  
# 获取页面对象内置的Session对象  
session = page1.session  
# 在第二个页面对象初始化时传递该对象  
page2 = SessionPage(session_or_options=session)
```

[上一页

🛩️ 概述](/SessionPage/intro)[下一页

🛩️ 访问网页](/SessionPage/visit)

* [✅️️ `SessionPage`初始化参数](#️️-sessionpage初始化参数)
* [✅️️ 直接创建](#️️-直接创建)
* [✅️️ 通过配置信息创建](#️️-通过配置信息创建)
  + [📌 使用方法](#-使用方法)
  + [📌 从指定 ini 文件创建](#-从指定-ini-文件创建)
  + [📌 不使用 ini 文件](#-不使用-ini-文件)
* [✅️️ 传递控制权](#️️-传递控制权)

* 🛫 SessionPage
* 🛩️ 查找元素

本页总览

🛩️ 查找元素
=======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

✅️️ 页面或元素内查找[​](#️️-页面或元素内查找 "✅️️ 页面或元素内查找的直接链接")
-------------------------------------------------

页面对象和元素对象都拥有`ele()`和`eles()`方法，用于获取其内部指定子元素。

### 📌 `ele()`[​](#-ele "-ele的直接链接")

用于查找其内部第一个条件匹配的元素。

页面对象和元素对象的`ele()`方法参数名称稍有不同，但用法一样。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` | 必填 | 元素的定位信息。可以是查询字符串，或 loc 元组 |
| `index` | `int` | `1` | 获取第几个匹配的元素，从`1`开始，可输入负数表示从后面开始数 |
| `timeout` | `float` | `None` | 此参数在这里没有作用 |

| 返回类型 | 说明 |
| --- | --- |
| `SessionElement` | 找到的第一个元素对象 |
| `NoneElement` | 未找到符合条件的元素时返回 |

---

### 📌 `eles()`[​](#-eles "-eles的直接链接")

此方法与`ele()`相似，但返回的是匹配到的所有元素组成的列表。

页面对象和元素对象都可调用这个方法。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` | 必填 | 元素的定位信息，可以是查询字符串，或 loc 元组 |
| `timeout` | `float` | `None` | 此参数在这里没有作用 |

| 返回类型 | 说明 |
| --- | --- |
| `SessionElementsList` | 元素对象组成的列表 |

---

✅️️ 相对定位[​](#️️-相对定位 "✅️️ 相对定位的直接链接")
-------------------------------------

以下方法可以以某元素为基准，在 DOM 中按照条件获取其直接子节点、同级节点、祖先元素、文档前后节点。

这里说的是 “节点”，不是 “元素”。因为相对定位可以获取除元素外的其它节点，包括文本、注释节点。

### 📌 获取父级元素[​](#-获取父级元素 "📌 获取父级元素的直接链接")

🔸 `parent()`

此方法获取当前元素某一级父元素，可指定筛选条件或层数。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `level_or_loc` | `int` `str` `Tuple[str, str]` | `1` | 第几级父元素，从`1`开始，或用定位符在祖先元素中进行筛选 |
| `index` | `int` | `1` | 当`level_or_loc`传入定位符，使用此参数选择第几个结果，从当前元素往上级数；当`level_or_loc`传入数字时，此参数无效 |
| `timeout` | `float` | `0` | 查找超时时间（秒） |

| 返回类型 | 说明 |
| --- | --- |
| `SessionElement` | 元素对象 |
| `NoneElement` | 未获取到结果时 |

---

### 📌 获取直接子节点[​](#-获取直接子节点 "📌 获取直接子节点的直接链接")

🔸 `child()`

此方法返回当前元素的一个直接子节点，可指定筛选条件和第几个。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` `int` | `''` | 用于筛选节点的查询语法，为`int`类型时`index`参数无效 |
| `index` | `int` | `1` | 查询结果中的第几个，从`1`开始，可输入负数表示倒数 |
| `timeout` | `float` | `None` | 无实际作用 |
| `ele_only` | `bool` | `True` | 是否只查找元素，为`False`时把文本、注释节点也纳入查找范围 |

| 返回类型 | 说明 |
| --- | --- |
| `str` | 获取非元素节点时返回字符串 |
| `SessionElement` | 元素对象 |
| `NoneElement` | 未获取到结果时 |

---

🔸 `children()`

此方法返回当前元素全部符合条件的直接子节点组成的列表，可用查询语法筛选。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` | `''` | 用于筛选节点的查询语法 |
| `timeout` | `float` | `None` | 无实际作用 |
| `ele_only` | `bool` | `True` | 是否只查找元素，为`False`时把文本、注释节点也纳入查找范围 |

| 返回类型 | 说明 |
| --- | --- |
| `SessionElementsList` | 结果元素对象组成的列表 |

---

### 📌 获取后面的同级节点[​](#-获取后面的同级节点 "📌 获取后面的同级节点的直接链接")

🔸 `next()`

此方法返回当前元素后面的某一个同级节点，可指定筛选条件和第几个。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` `int` | `''` | 用于筛选节点的查询语法，为`int`类型时`index`参数无效 |
| `index` | `int` | `1` | 查询结果中的第几个，从`1`开始，可输入负数表示倒数 |
| `timeout` | `float` | `None` | 无实际作用 |
| `ele_only` | `bool` | `True` | 是否只查找元素，为`False`时把文本、注释节点也纳入查找范围 |

| 返回类型 | 说明 |
| --- | --- |
| `str` | 获取非元素节点时返回字符串 |
| `SessionElement` | 元素对象 |
| `NoneElement` | 未获取到结果时 |

---

🔸 `nexts()`

此方法返回当前元素后面全部符合条件的同级节点组成的列表，可用查询语法筛选。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` | `''` | 用于筛选节点的查询语法 |
| `timeout` | `float` | `None` | 无实际作用 |
| `ele_only` | `bool` | `True` | 是否只查找元素，为`False`时把文本、注释节点也纳入查找范围 |

| 返回类型 | 说明 |
| --- | --- |
| `SessionElementsList` | 结果元素对象组成的列表 |

---

### 📌 获取前面的同级节点[​](#-获取前面的同级节点 "📌 获取前面的同级节点的直接链接")

🔸 `prev()`

此方法返回当前元素前面的某一个同级节点，可指定筛选条件和第几个。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` `int` | `''` | 用于筛选节点的查询语法，为`int`类型时`index`参数无效 |
| `index` | `int` | `1` | 查询结果中的第几个，从`1`开始，可输入负数表示倒数 |
| `timeout` | `float` | `None` | 无实际作用 |
| `ele_only` | `bool` | `True` | 是否只查找元素，为`False`时把文本、注释节点也纳入查找范围 |

| 返回类型 | 说明 |
| --- | --- |
| `str` | 获取非元素节点时返回字符串 |
| `SessionElement` | 元素对象 |
| `NoneElement` | 未获取到结果时 |

---

🔸 `prevs()`

此方法返回当前元素前面全部符合条件的同级节点组成的列表，可用查询语法筛选。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` | `''` | 用于筛选节点的查询语法 |
| `timeout` | `float` | `None` | 无实际作用 |
| `ele_only` | `bool` | `True` | 是否只查找元素，为`False`时把文本、注释节点也纳入查找范围 |

| 返回类型 | 说明 |
| --- | --- |
| `SessionElementsList` | 结果元素对象组成的列表 |

---

### 📌 在后面文档中查找节点[​](#-在后面文档中查找节点 "📌 在后面文档中查找节点的直接链接")

🔸 `after()`

此方法返回当前元素后面的某一个节点，可指定筛选条件和第几个。查找范围不限同级节点，而是整个 DOM 文档。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` `int` | `''` | 用于筛选节点的查询语法，为`int`类型时`index`参数无效 |
| `index` | `int` | `1` | 查询结果中的第几个，从`1`开始，可输入负数表示倒数 |
| `timeout` | `float` | `None` | 无实际作用 |
| `ele_only` | `bool` | `True` | 是否只查找元素，为`False`时把文本、注释节点也纳入查找范围 |

| 返回类型 | 说明 |
| --- | --- |
| `str` | 获取非元素节点时返回字符串 |
| `SessionElement` | 元素对象 |
| `NoneElement` | 未获取到结果时 |

---

🔸 `afters()`

此方法返回当前元素后面符合条件的全部节点组成的列表，可用查询语法筛选。查找范围不限同级节点，而是整个 DOM 文档。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` | `''` | 用于筛选节点的查询语法 |
| `timeout` | `float` | `None` | 无实际作用 |
| `ele_only` | `bool` | `True` | 是否只查找元素，为`False`时把文本、注释节点也纳入查找范围 |

| 返回类型 | 说明 |
| --- | --- |
| `SessionElementsList` | 结果元素对象组成的列表 |

---

### 📌 在前面文档中查找节点[​](#-在前面文档中查找节点 "📌 在前面文档中查找节点的直接链接")

🔸 `before()`

此方法返回当前元素前面的某一个符合条件的节点，可指定筛选条件和第几个。查找范围不限同级节点，而是整个 DOM 文档。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` `int` | `''` | 用于筛选节点的查询语法，为`int`类型时`index`参数无效 |
| `index` | `int` | `1` | 查询结果中的第几个，从`1`开始，可输入负数表示倒数 |
| `timeout` | `float` | `None` | 无实际作用 |
| `ele_only` | `bool` | `True` | 是否只查找元素，为`False`时把文本、注释节点也纳入查找范围 |

| 返回类型 | 说明 |
| --- | --- |
| `str` | 获取非元素节点时返回字符串 |
| `SessionElement` | 元素对象 |
| `NoneElement` | 未获取到结果时 |

---

🔸 `befores()`

此方法返回当前元素前面全部符合条件的节点组成的列表，可用查询语法筛选。查找范围不限同级节点，而是整个 DOM 文档。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locator` | `str` `Tuple[str, str]` | `''` | 用于筛选节点的查询语法 |
| `timeout` | `float` | `None` | 无实际作用 |
| `ele_only` | `bool` | `True` | 是否只查找元素，为`False`时把文  本、注释节点也纳入查找范围 |

| 返回类型 | 说明 |
| --- | --- |
| `SessionElementsList` | 结果元素对象组成的列表 |

---

  

✅️️ 同时匹配多个定位符[​](#️️-同时匹配多个定位符 "✅️️ 同时匹配多个定位符的直接链接")
----------------------------------------------------

所有页面或元素对象都有`find()`方法，可接收多个定位符，同时查找多个（批）不同定位符的元素。

以`dict`方法返回每个定位符结果。

说明

当`first_ele`为`True`时，如果一个定位符没有被执行过查找，它返回的结果为`None`。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `locators` | `List[str]` `Tuple[str]` `str` `tuple` | 必填 | 定位符组成的列表 |
| `any_one` | `bool` | `True` | 是否任何一个定位符找到结果即返回 |
| `first_ele` | `bool` | `True` | 每个定位符获取第一个元素还是所有元素 |
| `timeout` | `float` | `None` | 超时时间（秒），为`None`使用该对象默认设置 |

说明

以下所说的 “定位符”，是`str`或`tuple`类型的。
“元素对象”，是`SessionElement`类型的，没有找到时是`NoneElement`类型的。
“元素对象组成的列表” 是`SessionElementsList`类型的。
`any_one`参数为`True`时，以`tuple`方式返回  找到目标的定位符和结果，为`False`时以`dict`方法返回每个定位符结果。

| 返回类型 | `any_one`参数取值 | 说明 |
| --- | --- | --- |
| `tuple(定位符, 元素对象)` | `True` | `first_ele`为`True`时，返回第一个有结果的定位符找到的第一个元素对象 |
| `tuple(定位符, 元素对象组成的列表)` | `True` | `first_ele`为`False`时，返回第一个有结果的定位符找到的所有元素对象 |
| `tuple(None, None)` | `True` | 所有定位符都没有找到元素，返回`(None, None)` |
| `dict{定位符: 元素对象}` | `False` | `first_ele`为`True`时，每个定位符返回第一个元素，找不到时为`NoneElement` |
| `dict{定位符: 元素对象组成的列表}` | `False` | `first_ele`为`False`时，每个定位符返回所有结果元素组成的列表 |

[上一页

🛩️ 获取页面信息](/SessionPage/get_page_info)[下一页

🛩️ 获取元素信息](/SessionPage/get_ele_info)

* [✅️️ 页面或元素内查找](#️️-页面或元素内查找)
  + [📌 `ele()`](#-ele)
  + [📌 `eles()`](#-eles)
* [✅️️ 相对定位](#️️-相对定位)
  + [📌 获取父级元素](#-获取父级元素)
  + [📌 获取直接子节点](#-获取直接子节点)
  + [📌 获取后面的同级节点](#-获取后面的同级节点)
  + [📌 获取前面的同级节点](#-获取前面的同级节点)
  + [📌 在后面文档中查找节点](#-在后面文档中查找节点)
  + [📌 在前面文档中查找节点](#-在前面文档中查找节点)
* [✅️️ 同时匹配多个定位符](#️️-同时匹配多个定位符)

* 🛫 SessionPage
* 🛩️ 获取元素信息

本页总览

🛩️ 获取元素信息
=========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

`SessionPage`对象获取的元素是`SessionElement`，本节介绍其属性。

假设`ele`为以下`div`元素的对象，本节示例均使用该元素：

```
<div id="div1" class="divs">Hello World!  
    <p>行元素</p>  
    <!--这是注释-->  
</div>
```

✅️️ `html`[​](#️️-html "️️-html的直接链接")
--------------------------------------

此属性返回元素的`outerHTML`文本。

**返回类型：**`str`

```
print(ele.html)
```

**输出：**

```
<div id="div1" class="divs">Hello World!  
    <p>行元素</p>  
    <!--这是注释-->  
</div>
```

---

✅️️ `inner_html`[​](#️️-inner_html "️️-inner_html的直接链接")
--------------------------------------------------------

此属性返回元素的`innerHTML`文本。

**返回类型：**`str`

```
print(ele.inner_html)
```

**输出：**

```
Hello World!  
    <p>行元素</p>  
    <!--这是注释-->
```

---

✅️️ `tag`[​](#️️-tag "️️-tag的直接链接")
-----------------------------------

此属性返回元素的标签名。

**返回类型：**`str`

```
print(ele.tag)
```

**输出：**

```
div
```

---

✅️️ `text`[​](#️️-text "️️-text的直接链接")
--------------------------------------

此属性返回元素内所有文本组合成的字符串。  
该字符串已格式化，即已转码，已去除多余换行符，符合人读取习惯，便于直接使用。

**返回类型：**`str`

```
print(ele.text)
```

**输出：**

```
Hello World!  
行元素
```

---

✅️️ `raw_text`[​](#️️-raw_text "️️-raw_text的直接链接")
--------------------------------------------------

此属性返回元素内原始文本。

**返回类型：**`str`

```
print(ele.raw_text)
```

输出（注意保留了元素间的空格和换行）：

```
Hello World!  
    行元素  
　      
　
```

---

  

✅️️ `texts()`[​](#️️-texts "️️-texts的直接链接")
-------------------------------------------

此方法返回元素内所有**直接**子节点的文本，包括元素和文本节点。 它有一个参数`text_node_only`，为`True`时则只获取只返回不被包裹的文本节点。这个方法适用于获取文本节点和元素节点混排的情况。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `text_node_only` | `bool` | `False` | 是否只返回文本节点 |

| 返回类型 | 说明 |
| --- | --- |
| `List[str]` | 文本列表 |

**示例：**

```
print(e.texts())    
print(e.texts(text_node_only=True))
```

**输出：**

```
['Hello World!', '行元素']  
['Hello World!']
```

---

✅️️ `comments`[​](#️️-comments "️️-comments的直接链接")
--------------------------------------------------

此属性以列表形式返回元素内的注释。

**返回类型：**`List[str]`

```
print(ele.comments)
```

**输出：**

```
[<!--这是注释-->]
```

---

✅️️ `attrs`[​](#️️-attrs "️️-attrs的直接链接")
-----------------------------------------

此属性以字典形式返回元素所有属性及值。

**返回类型：**`dict`

```
print(ele.attrs)
```

**输出：**

```
{'id': 'div1', 'class': 'divs'}
```

---

✅️️ `attr()`[​](#️️-attr "️️-attr的直接链接")
----------------------------------------

此方法返回元素某个 attribute 属性值。它接收一个字符串参数，返回该属性值文本，无该属性时返回`None`。  
此属性返回的`src`、`href`属性为已补充完整的路径。`text`属性为已格式化文本。
如果要获取未补充完整路径的`src`或`href`属性，可以用`attrs['src']`。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `name` | `str` | 必填 | 属性名称 |

| 返回类型 | 说明 |
| --- | --- |
| `str` | 属性值文本 |
| `None` | 没有该属性返回`None` |

**示例：**

```
print(ele.attr('id'))
```

**输出：**

```
div1
```

---

✅️️ `link`[​](#️️-link "️️-link的直接链接")
--------------------------------------

此方法返回元素的 href 属性或 src 属性，没有这两个属性则返回`None`。

**返回类型：**`str`

```
<a href='http://www.baidu.com'>百度</a>
```

假设`a_ele`为以上元素的对象：

```
print(a_ele.link)
```

**输出：**

```
http://www.baidu.com
```

---

### 📌 `child_count`[​](#-child_count "-child_count的直接链接")

此属性返回元素内第一级子元素个数。

**类型：**`int`

---

✅️️ `page`[​](#️️-page "️️-page的直接链接")
--------------------------------------

此属性返回元素所在的页面对象。由 html 文本直接生成的`SessionElement`的`page`属性为`None`。

**返回类型：**`SessionPage`、`WebPage`

```
page = ele.page
```

---

✅️️ `xpath`[​](#️️-xpath "️️-xpath的直接链接")
-----------------------------------------

此属性返回当前元素在页面中 xpath 的绝对路径。

**返回类型：**`str`

```
print(ele.xpath)
```

**输出：**

```
/html/body/div
```

---

✅️️ `css_path`[​](#️️-css_path "️️-css_path的直接链接")
--------------------------------------------------

此属性返回当前元素在页面中 css selector 的绝对路径。

**返回类型：**`str`

```
print(ele.css_path)
```

**输出：**

```
:nth-child(1)>:nth-child(1)>:nth-child(1)
```

---

✅️️ 元素列表中批量获取信息[​](#️️-元素列表中批量获取信息 "✅️️ 元素列表中批量获取信息的直接链接")
----------------------------------------------------------

`eles()`等返回的元素列表，自带`get`属性，可用于获取指定信息。

### 📌 示例[​](#-示例 "📌 示例的直接链接")

```
from DrissionPage import SessionPage  
  
page = SessionPage()  
page.get('https://www.baidu.com')  
eles = page('#s-top-left').eles('t:a')  
print(eles.get.texts())  # 获取所有元素的文本
```

**输出：**

```
['新闻', 'hao123', '地图', '贴吧', '视频', '图片', '网盘', '文库', '更多', '翻译', '学术', '百科', '知道', '健康', '营销推广', '直播', '音乐', '橙篇', '查看全部百度产品 >']
```

### 📌 `get.attrs()`[​](#-getattrs "-getattrs的直接链接")

此方法用于返回所有元素指定的 attribute 属性组成的列表。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `name` | `str` | 必填 | 属性名称 |

| 返回类型 | 说明 |
| --- | --- |
| `List[str]` | 属性文本组成的列表 |

---

### 📌 `get.links()`[​](#-getlinks "-getlinks的直接链接")

此方法用于返回所有元素的`link`属性组成的列表。

**参数：** 无

| 返回类型 | 说明 |
| --- | --- |
| `List[str]` | 链接文本组成的列表 |

---

### 📌 `get.texts()`[​](#-gettexts "-gettexts的直接链接")

此方法用于返回所有元素的`text`属性组成的列表。

**参数：** 无

| 返回类型 | 说明 |
| --- | --- |
| `List[str]` | 元素文本组成的列表 |

---

✅️️ 实际示例[​](#️️-实际示例 "✅️️ 实际示例的直接链接")
-------------------------------------

以下示例可直接运行查看结果：

```
from DrissionPage import SessionPage  
  
page = SessionPage()  
page.get('https://gitee.com/explore')  
  
# 获取推荐目录下所有 a 元素  
li_eles = page('tag:ul@text():全部推荐项目').eles('t:a')  
  
# 遍历列表  
for i in li_eles:    
    # 获取并打印标签名、文本、href 属性  
    print(i.tag, i.text, i.attribute('href'))
```

**输出：**

```
a 全部推荐项目 https://gitee.com/explore/all  
a 前沿技术 https://gitee.com/explore/new-tech  
a 智能硬件 https://gitee.com/explore/hardware  
a IOT/物联网/边缘计算 https://gitee.com/explore/iot  
a 车载应用 https://gitee.com/explore/vehicle  
以下省略……
```

[上一页

🛩️ 查找元素](/SessionPage/get_ele)[下一页

🛩️ 页面设置](/SessionPage/settings)

* [✅️️ `html`](#️️-html)
* [✅️️ `inner_html`](#️️-inner_html)
* [✅️️ `tag`](#️️-tag)
* [✅️️ `text`](#️️-text)
* [✅️️ `raw_text`](#️️-raw_text)
* [✅️️ `texts()`](#️️-texts)
* [✅️️ `comments`](#️️-comments)
* [✅️️ `attrs`](#️️-attrs)
* [✅️️ `attr()`](#️️-attr)
* [✅️️ `link`](#️️-link)
  + [📌 `child_count`](#-child_count)
* [✅️️ `page`](#️️-page)
* [✅️️ `xpath`](#️️-xpath)
* [✅️️ `css_path`](#️️-css_path)
* [✅️️ 元素列表中批量获取信息](#️️-元素列表中批量获取信息)
  + [📌 示例](#-示例)
  + [📌 `get.attrs()`](#-getattrs)
  + [📌 `get.links()`](#-getlinks)
  + [📌 `get.texts()`](#-gettexts)
* [✅️️ 实际示例](#️️-实际示例)

* 🛫 SessionPage
* 🛩️ 获取页面信息

本页总览

🛩️ 获取页面信息
=========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

成功访问网页后，可使用`SessionPage`对象自身属性和方法获取页面信息。

```
from DrissionPage import SessionPage  
  
page = SessionPage()  
page.get('http://www.baidu.com')  
# 获取页面标题  
print(page.title)  
# 获取页面html  
print(page.html)
```

**输出：**

```
百度一下，你就知道  
<!DOCTYPE html>  
<!--STATUS OK--><html> <head><meta http-equi...
```

---

✅️️ 页面信息[​](#️️-页面信息 "✅️️ 页面信息的直接链接")
-------------------------------------

### 📌 `url`[​](#-url "-url的直接链接")

此属性返回当前访问的 url。

**类型：**`str`

---

### 📌 `url_available`[​](#-url_available "-url_available的直接链接")

此属性以布尔值返回当前链接是否可用。

**类型：**`bool`

---

### 📌 `title`[​](#-title "-title的直接链接")

此属性返回当前页面`title`文本。

**类型：**`str`

---

### 📌 `raw_data`[​](#-raw_data "-raw_data的直接链接")

此属性返回访问到的元素数据，即`Response`对象的`content`属性。

**类型：**`bytes`

---

### 📌 `html`[​](#-html "-html的直接链接")

此属性返回当前页面 html 文本。

**类型：**`str`

---

### 📌 `json`[​](#-json "-json的直接链接")

此属性把返回内容解析成 json。  
比如请求接口时，若返回内容是 json 格式，用`html`属性获取的话会得到一个字符串，用此属性获取可将其解析成`dict`。
支持访问 `*.json` 文件，也支持 API 返回的json字符串。

**类型：**`dict`

---

### 📌 `user_agent`[​](#-user_agent "-user_agent的直接链接")

此属性返回当前页面 user\_agent 信息。

**类型：**`str`

---

✅️️ 运行参数信息[​](#️️-运行参数信息 "✅️️ 运行参数信息的直接链接")
-------------------------------------------

### 📌 `timeout`[​](#-timeout "-timeout的直接链接")

此属性返回网络请求超时时间，默认为 10 秒。

**类型：**`int`、`float`

---

### 📌 `retry_times`[​](#-retry_times "-retry_times的直接链接")

此属性为网络连接失败时的重试次数，默认为`3`。

**类型：**`int`

---

### 📌 `retry_interval`[​](#-retry_interval "-retry_interval的直接链接")

此属性为网络连接失败时的重试等待间隔秒数，默认为`2`。

**类型：**`int`、`float`

---

### 📌 `encoding`[​](#-encoding "-encoding的直接链接")

此属性返回用户主动设置的编码格式。

---

  

✅️️ cookies 信息[​](#️️-cookies-信息 "✅️️ cookies 信息的直接链接")
-------------------------------------------------------

### 📌 `cookies()`[​](#-cookies "-cookies的直接链接")

此方法返回 cookies 信息。

**类型：**`dict`、`list`

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `all_domains` | `bool` | `False` | 是否返回所有 cookies，为`False`只返回当前 url 的 |
| `all_info` | `bool` | `False` | 返回的 cookies 是否包含所有信息，`False`时只包含`name`、`value`、`domain`信息 |

| 返回类型 | 说明 |
| --- | --- |
| `CookiesList` | cookies 组成的列表 |

`cookies()`方法返回的列表可转换为其它指定格式。

* `cookies().as_str()`：`'name1=value1; name2=value2'`格式的字符串
* `cookies().as_dict()`：`{name1: value1, name2: value2}`格式的字典
* `cookies().as_json()`：json 格式的字符串

说明

`as_str()`和`as_dict()`都只会保留`'name'`和`'value'`字段。

**示例：**

```
from DrissionPage import SessionPage  
  
page = SessionPage()  
page.get('http://www.baidu.com')  
page.get('http://gitee.com')  
  
for i in page.cookies(all_domains=True):  
    print(i)
```

**输出：**

```
{'name': 'BDORZ', 'value': '27875', 'domain': '.baidu.com'}  
{'name': 'BEC', 'value': '1f1759dfh65j65j5j4feb0357', 'domain': 'gitee.com'}
```

---

✅️️ 内嵌对象[​](#️️-内嵌对象 "✅️️ 内嵌对象的直接链接")
-------------------------------------

### 📌 `session`[​](#-session "-session的直接链接")

此属性返回当前页面对象使用的`Session`对象。

**类型：**`Session`

---

### 📌 `response`[​](#-response "-response的直接链接")

此属性为请求网页后生成的`Response`对象，本库没实现的功能可直接获取此属性调用 requests 库的原生功能。

**类型：**`Response`

```
# 打印连接状态  
r = page.response  
print(r.status_code)
```

[上一页

🛩️ 访问网页](/SessionPage/visit)[下一页

🛩️ 查找元素](/SessionPage/get_ele)

* [✅️️ 页面信息](#️️-页面信息)
  + [📌 `url`](#-url)
  + [📌 `url_available`](#-url_available)
  + [📌 `title`](#-title)
  + [📌 `raw_data`](#-raw_data)
  + [📌 `html`](#-html)
  + [📌 `json`](#-json)
  + [📌 `user_agent`](#-user_agent)
* [✅️️ 运行参数信息](#️️-运行参数信息)
  + [📌 `timeout`](#-timeout)
  + [📌 `retry_times`](#-retry_times)
  + [📌 `retry_interval`](#-retry_interval)
  + [📌 `encoding`](#-encoding)
* [✅️️ cookies 信息](#️️-cookies-信息)
  + [📌 `cookies()`](#-cookies)
* [✅️️ 内嵌对象](#️️-内嵌对象)
  + [📌 `session`](#-session)
  + [📌 `response`](#-response)

* 🛫 SessionPage
* 🛩️ 概述

🛩️ 概述
=====

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

`SessionPage`对象和`WebPage`对象的 s 模式，可用收发数据包的形式访问网页。

`SessionPage`是一个使用使用`Session`（requests 库）对象的页面，封装了网络连接和结果解析功能，使收发数据包也可以像操作页面一样便利。

**示例：**

获取 gitee 推荐项目第一页所有项目。

```
# 导入  
from DrissionPage import SessionPage  
# 创建页面对象  
page = SessionPage()  
# 访问网页  
page.get('https://gitee.com/explore/all')  
# 在页面中查找元素  
items = page.eles('t:h3')  
# 遍历元素  
for item in items[:-1]:  
    # 获取当前<h3>元素下的<a>元素  
    lnk = item('tag:a')  
    # 打印<a>元素文本和href属性  
    print(lnk.text, lnk.link)
```

**输出：**

```
七年觐汐/wx-calendar https://gitee.com/qq_connect-EC6BCC0B556624342/wx-calendar  
ThingsPanel/thingspanel-go https://gitee.com/ThingsPanel/thingspanel-go  
APITable/APITable https://gitee.com/apitable/APITable  
Indexea/ideaseg https://gitee.com/indexea/ideaseg  
CcSimple/vue-plugin-hiprint https://gitee.com/CcSimple/vue-plugin-hiprint  
william_lzw/ExDUIR.NET https://gitee.com/william_lzw/ExDUIR.NET  
anolis/ancert https://gitee.com/anolis/ancert  
cozodb/cozo https://gitee.com/cozodb/cozo  
后面省略...
```

[上一页

🛰️ Page 对象](/browser_control/pages)[下一页

🛩️ 创建页面对象](/SessionPage/create_obj)

* 🛫 SessionPage
* 🛩️ 启动配置

本页总览

🛩️ 启动配置
=======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

我们用`SessionOptions`对象管理`SessionPage`对象初始配置。

注意

`SessionOptions`仅用于管理启动配置，程序启动后再修改无效。

✅️️ 创建对象[​](#️️-创建对象 "✅️️ 创建对象的直接链接")
-------------------------------------

### 📌 导入[​](#-导入 "📌 导入的直接链接")

```
from DrissionPage import SessionOptions
```

---

### 📌 初始化参数[​](#-初始化参数 "📌 初始化参数的直接链接")

`SessionOptions`对象用于管理`Session`对象的初始化配置。可从配置文件中读取配置来进行初始化。

| 初始化参数 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `read_file` | `bool` | `True` | 是否从 ini 文件中读取配置信息，为`False`则用默认配置创建 |
| `ini_path` | `Path` `str` | `None` | 指定 ini 文件路径，为`None`则读取内置 ini 文件 |

创建配置对象：

```
from DrissionPage import SessionOptions  
  
so = SessionOptions()
```

默认情况下，`SessionOptions`对象会从 ini 文件中读取配置信息，当指定`read_file`参数为`False`时，则以默认配置创建。

提醒

对象创建时已带有默认 headers，如要清除，可调用`clear_headers()`方法。

---

✅️️ 使用方法[​](#️️-使用方法 "✅️️ 使用方法的直接链接")
-------------------------------------

创建配置对象后，可调整配置内容，然后在页面对象创建时以参数形式把配置对象传递进去。

```
from DrissionPage import SessionPage, SessionOptions  
  
# 创建配置对象（默认从 ini 文件中读取配置）  
so = SessionOptions()  
# 设置代理  
so.set_proxies('http://localhost:1080')  
# 设置 cookies  
cookies = ['key1=val1; domain=****', 'key2=val2; domain=****']  
so.set_cookies(cookies)  
  
# 以该配置创建页面对象  
page = SessionPage(session_or_options=so)
```

---

  

✅️️ 用于设置的方法[​](#️️-用于设置的方法 "✅️️ 用于设置的方法的直接链接")
----------------------------------------------

### 📌 `set_headers()`[​](#-set_headers "-set_headers的直接链接")

该方法用于设置整个 headers 参数，传入值会覆盖原来的 headers。

headers 可以是`dict`格式的，也可以是文本格式。

文本格式不同字段用`\n`分隔，字段 key 和 value 用`': '`分隔，即从浏览器直接复制的格式。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `headers` | `dict` `str` | 必填 | headers 信息 |

| 返回类型 | 说明 |
| --- | --- |
| `SessionOptions` | 配置对象本身 |

**示例：**

```
so.set_headers = {'user-agent': 'Mozilla/5.0 (Macint...', 'connection': 'keep-alive' ...}
```

---

### 📌 `set_a_header()`[​](#-set_a_header "-set_a_header的直接链接")

该方法用于设置`headers`中的一个项。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `name` | `str` | 必填 | 设置名称 |
| `value` | `str` | 必填 | 设置值 |

| 返回类型 | 说明 |
| --- | --- |
| `SessionOptions` | 配置对象本身 |

**示例：**

```
so.set_a_header('accept', 'text/html')  
so.set_a_header('Accept-Charset', 'GB2312')
```

**输出：**

```
{'accept': 'text/html', 'accept-charset': 'GB2312'}
```

---

### 📌 `remove_a_header()`[​](#-remove_a_header "-remove_a_header的直接链接")

此方法用于从`headers`中移除一个设置项。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `name` | `str` | 必填 | 要删除的设置 |

| 返回类型 | 说明 |
| --- | --- |
| `SessionOptions` | 配置对象本身 |

**示例：**

```
so.remove_a_header('accept')
```

---

### 📌 `clear_headers()`[​](#-clear_headers "-clear_headers的直接链接")

此方法用于清空已设置的`headers`参数。

**参数：** 无

| 返回类型 | 说明 |
| --- | --- |
| `SessionOptions` | 配置对象自身 |

---

### 📌 `set_cookies()`[​](#-set_cookies "-set_cookies的直接链接")

此方法用于设置一个或多个 cookie，每次设置会覆盖之前所有 cookies 信息。

详细用法见实用教程相关章节。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `cookies` | `Cookie` `CookieJar` `list` `tuple` `str` `dict` | 必填 | cookies |

| 返回类型 | 说明 |
| --- | --- |
| `SessionOptions` | 配置对象本身 |

**示例：**

```
cookies = ['key1=val1; domain=****', 'key2=val2; domain=****']  
so.set_cookies(cookies)
```

---

### 📌 `set_timeout()`[​](#-set_timeout "-set_timeout的直接链接")

此方法用于设置连接超时属性。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `second` | `float` | 必填 | 连接等待秒数 |

| 返回类型 | 说明 |
| --- | --- |
| `SessionOptions` | 配置对象本身 |

---

### 📌 `set_retry()`[​](#-set_retry "-set_retry的直接链接")

此方法用于设置页面连接超时时的重试次数和间隔。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `times` | `int` | `None` | 连接失败重试次数 |
| `interval` | `float` | `None` | 连接失败重试间隔（秒） |

| 返回类型 | 说明 |
| --- | --- |
| `SessionOptions` | 配置对象本身 |

---

### 📌 `retry_times`[​](#-retry_times "-retry_times的直接链接")

该属性返回连接失败时的重试次数。

**类型：**`int`

---

### 📌 `retry_interval`[​](#-retry_interval "-retry_interval的直接链接")

该属性返回连接失败时的重试间隔（秒）。

**类型：**`float`

---

### 📌 `set_proxies()`[​](#-set_proxies "-set_proxies的直接链接")

此方法用于设置代理信息。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `http` | `str` | `None` | http 代理地址 |
| `https` | `str` | `None` | https 代理地址 |

| 返回类型 | 说明 |
| --- | --- |
| `SessionOptions` | 配置对象本身 |

**示例：**

```
so.set_proxies('http://127.0.0.1:1080')
```

---

### 📌 `set_download_path()`[​](#-set_download_path "-set_download_path的直接链接")

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `path` | `str` `Path` | 必填 | 默认下载保存路径 |

| 返回类型 | 说明 |
| --- | --- |
| `SessionOptions` | 配置对象本身 |

---

### 📌 `set_auth()`[​](#-set_auth "-set_auth的直接链接")

此方法用于设置认证元组信息。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `auth` | `tuple` `HTTPBasicAuth` | 必填 | 认证元组或对象 |

| 返回类型 | 说明 |
| --- | --- |
| `SessionOptions` | 配置对象本身 |

---

### 📌 `set_hooks()`[​](#-set_hooks "-set_hooks的直接链接")

此方法用于设置回调方法。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `hooks` | `dict` | 必填 | 回调方法 |

| 返回类型 | 说明 |
| --- | --- |
| `SessionOptions` | 配置对象本身 |

---

### 📌 `set_params()`[​](#-set_params "-set_params的直接链接")

此方法用于设置查询参数。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `params` | `dict` | 必填 | 查询参数字典 |

| 返回类型 | 说明 |
| --- | --- |
| `SessionOptions` | 配置对象本身 |

---

### 📌 `set_cert()`[​](#-set_cert "-set_cert的直接链接")

此方法用于设置 SSL 客户端证书文件的路径（.pem格式），或 ('cert', 'key') 元组。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `cert` | `str` `tuple` | 必填 | 证书路径或元组 |

| 返回类型 | 说明 |
| --- | --- |
| `SessionOptions` | 配置对象本身 |

---

### 📌 `set_verify()`[​](#-set_verify "-set_verify的直接链接")

此方法用于设置是否验证SSL证书。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_off` | `bool` | 必填 | `bool`表示开或关 |

| 返回类型 | 说明 |
| --- | --- |
| `SessionOptions` | 配置对象本身 |

---

### 📌 `add_adapter()`[​](#-add_adapter "-add_adapter的直接链接")

此方法用于添加适配器。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `url` | `str` | 必填 | 适配器对应 url |
| `adapter` | `HTTPAdapter` | 必填 | 适配器对象 |

| 返回类型 | 说明 |
| --- | --- |
| `SessionOptions` | 配置对象本身 |

---

### 📌 `set_stream()`[​](#-set_stream "-set_stream的直接链接")

此方法用于设置是否使用流式响应内容。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_off` | `bool` | 必填 | `bool`表示开或关 |

| 返回类型 | 说明 |
| --- | --- |
| `SessionOptions` | 配置对象本身 |

---

### 📌 `set_trust_env()`[​](#-set_trust_env "-set_trust_env的直接链接")

此方法用于设置是否信任环境。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_off` | `bool` | 必填 | `bool`表示开或关 |

| 返回类型 | 说明 |
| --- | --- |
| `SessionOptions` | 配置对象本身 |

---

### 📌 `set_max_redirects()`[​](#-set_max_redirects "-set_max_redirects的直接链接")

此方法用于设置最大重定向次数。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `times` | `int` | 必填 | 最大重定向次数 |

| 返回类型 | 说明 |
| --- | --- |
| `SessionOptions` | 配置对象本身 |

---

✅️️ 保存设置到文件[​](#️️-保存设置到文件 "✅️️ 保存设置到文件的直接链接")
----------------------------------------------

您可以把不同的配置保存到各自的 ini 文件，以便适应不同的场景。

注意

`hooks`和`adapters`配置是不会保存到文件中的。

### 📌 `save()`[​](#-save "-save的直接链接")

此方法用于保存配置项到一个 ini 文件。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `path` | `str` `Path` | `None` | ini 文件的路径， 传入`None`保存到当前读取的配置文件 |

| 返回类型 | 说明 |
| --- | --- |
| `str` | 保存的 ini 文件绝对路径 |

**示例：**

```
# 保存当前读取的ini文件  
so.save()  
  
# 把当前配置保存到指定的路径  
so.save(path=r'D:\tmp\settings.ini')
```

---

### 📌 `save_to_default()`[​](#-save_to_default "-save_to_default的直接链接")

此方法用于保存配置项到固定的默认 ini 文件。默认 ini 文件是指随 DrissionPage 内置的那个。

**参数：** 无

| 返回类型 | 说明 |
| --- | --- |
| `str` | 保存的 ini 文件绝对路径 |

**示例：**

```
so.save_to_default()
```

---

✅️️ `SessionOptions`属性[​](#️️-sessionoptions属性 "️️-sessionoptions属性的直接链接")
--------------------------------------------------------------------------

### 📌 `headers`[​](#-headers "-headers的直接链接")

该属性返回 headers 设置信息。

**类型：**`dict`

---

### 📌 `cookies`[​](#-cookies "-cookies的直接链接")

此属性以`list`方式返回 cookies 设置信息。

**类型：**`list`

---

### 📌 `proxies`[​](#-proxies "-proxies的直接链接")

此属性返回代理信息。

**类型：**`dict`
**格式：**`{'http': 'http://**.**.**.**:****', 'https': 'http://**.**.**.**:****'}`

---

### 📌 `auth`[​](#-auth "-auth的直接链接")

此属性返回认证设置。

**类型：**`tuple`、`HTTPBasicAuth`

---

### 📌 `hooks`[​](#-hooks "-hooks的直接链接")

此属性返回回调方法设置。

**类型：**`dict`

---

### 📌 `params`[​](#-params "-params的直接链接")

此属性返回查询参数设置。

**类型：**`dict`

---

### 📌 `verify`[​](#-verify "-verify的直接链接")

此属性返回是否验证 SSL 证书设置。

**类型：**`bool`

---

### 📌 `cert`[​](#-cert "-cert的直接链接")

此属性返回 SSL 证书设置。

**类型：**`str`、`tuple`

---

### 📌 `adapters`[​](#-adapters "-adapters的直接链接")

此属性返回适配器设置。

**类型：**`List[HTTPAdapter]`

---

### 📌 `stream`[​](#-stream "-stream的直接链接")

此属性返回是否使用流式响应设置。

**类型：**`bool`

---

### 📌 `trust_env`[​](#-trust_env "-trust_env的直接链接")

此属性返回是否信任环境设置。

**类型：**`bool`

---

### 📌 `max_redirects`[​](#-max_redirects "-max_redirects的直接链接")

此属性返回`max_redirects`设置。

**类型：**`int`

---

### 📌 `timeout`[​](#-timeout "-timeout的直接链接")

此属性返回连接超时设置。

**类型：**`int`、`float`

---

### 📌 `download_path`[​](#-download_path "-download_path的直接链接")

此属性返回默认下载路径设置。

**类型：**`str`

[上一页

🛩️ 页面设置](/SessionPage/settings)[下一页

⤵️ 概述](/download/intro)

* [✅️️ 创建对象](#️️-创建对象)
  + [📌 导入](#-导入)
  + [📌 初始化参数](#-初始化参数)
* [✅️️ 使用方法](#️️-使用方法)
* [✅️️ 用于设置的方法](#️️-用于设置的方法)
  + [📌 `set_headers()`](#-set_headers)
  + [📌 `set_a_header()`](#-set_a_header)
  + [📌 `remove_a_header()`](#-remove_a_header)
  + [📌 `clear_headers()`](#-clear_headers)
  + [📌 `set_cookies()`](#-set_cookies)
  + [📌 `set_timeout()`](#-set_timeout)
  + [📌 `set_retry()`](#-set_retry)
  + [📌 `retry_times`](#-retry_times)
  + [📌 `retry_interval`](#-retry_interval)
  + [📌 `set_proxies()`](#-set_proxies)
  + [📌 `set_download_path()`](#-set_download_path)
  + [📌 `set_auth()`](#-set_auth)
  + [📌 `set_hooks()`](#-set_hooks)
  + [📌 `set_params()`](#-set_params)
  + [📌 `set_cert()`](#-set_cert)
  + [📌 `set_verify()`](#-set_verify)
  + [📌 `add_adapter()`](#-add_adapter)
  + [📌 `set_stream()`](#-set_stream)
  + [📌 `set_trust_env()`](#-set_trust_env)
  + [📌 `set_max_redirects()`](#-set_max_redirects)
* [✅️️ 保存设置到文件](#️️-保存设置到文件)
  + [📌 `save()`](#-save)
  + [📌 `save_to_default()`](#-save_to_default)
* [✅️️ `SessionOptions`属性](#️️-sessionoptions属性)
  + [📌 `headers`](#-headers)
  + [📌 `cookies`](#-cookies)
  + [📌 `proxies`](#-proxies)
  + [📌 `auth`](#-auth)
  + [📌 `hooks`](#-hooks)
  + [📌 `params`](#-params)
  + [📌 `verify`](#-verify)
  + [📌 `cert`](#-cert)
  + [📌 `adapters`](#-adapters)
  + [📌 `stream`](#-stream)
  + [📌 `trust_env`](#-trust_env)
  + [📌 `max_redirects`](#-max_redirects)
  + [📌 `timeout`](#-timeout)
  + [📌 `download_path`](#-download_path)

* 🛫 SessionPage
* 🛩️ 页面设置

本页总览

🛩️ 页面设置
=======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

本节介绍`SessionPage`运行参数设置。

这些设置是全局参数，设置后每次请求都会使用它们。

**示例：**

```
from DrissionPage import SessionPage  
  
page = SessionPage()  
page.set.cookies([{'name': 'a', 'value': '1'}, {'name': 'b', 'value': '2'}])
```

✅️️ `set.retry_times()`[​](#️️-setretry_times "️️-setretry_times的直接链接")
-----------------------------------------------------------------------

此方法用于设置连接失败时重连次数。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `times` | `int` | 必填 | 次数 |

**返回：**`None`

✅️️ `set.retry_interval()`[​](#️️-setretry_interval "️️-setretry_interval的直接链接")
--------------------------------------------------------------------------------

此方法用于设置连接失败时重连间隔。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `interval` | `float` | 必填 | 秒数 |

**返回：**`None`

✅️️ `set.timeout()`[​](#️️-settimeout "️️-settimeout的直接链接")
-----------------------------------------------------------

此方法用于设置连接超时时间（秒）。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `second` | `float` | 必填 | 秒数 |

**返回：**`None`

**示例：**

```
page.set.timeout(20)
```

---

✅️️ `set.encoding()`[​](#️️-setencoding "️️-setencoding的直接链接")
--------------------------------------------------------------

此方法用于设置网页编码。

默认情况下，程序会自动从 headers、页面上获取编码，但总有些奇葩网页的编码不准确。这时候可以主动设置编码。

可以针对已获取的`Rsponse`对象设置，或作为整体设置对之后的连接都有效。

| 参数 名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `encoding` | `str` | 必填 | 编码名称，如果要取消之前的设置，传入`None` |
| `set_all` | `bool` | `True` | 是否设置对象参数，为`False`则只设置当前`Response`对象 |

**返回：**`None`

---

✅️️ `set.cookies()`[​](#️️-setcookies "️️-setcookies的直接链接")
-----------------------------------------------------------

此方法用于设置一个或多个 cookie。

设置一个 cookie 支持的格式：

* `Cookie`：单个`Cookie`对象
* `str`：`'name=value; domain=****; ...'`或`'name=****; value=****; domain=****; ...'`格式，只支持用`';'`分隔
* `dict`：`{'name': '****', 'value': '****', 'domain': '****', ...}`或`{name: value, 'domain': '****', ...}`格式

设置多个 cookie 支持的格式：

* `list`或`tuple`：上面几种形式的单个 cookie 放到列表中传入即可
* `dict`：`{name1: value1, name2: value2, ..., 'domain': '****', ...}`格式
* `str`：`'name1=value1; name2=value2; ... domain=****; ...'`格式，多个 cookie 之间只能用`';'`分隔
* `CookieJar`：单个`CookieJar`对象

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `cookies` | `Cookie` `CookieJar` `list` `tuple` `str` `dict` | 必填 | cookies 信息 |

**返回：**`None`

---

  

✅️️ `set.cookies.clear()`[​](#️️-setcookiesclear "️️-setcookiesclear的直接链接")
---------------------------------------------------------------------------

此方法用于清除所有 cookie。

**参数：** 无

**返回：**`None`

---

✅️️ `set.cookies.remove()`[​](#️️-setcookiesremove "️️-setcookiesremove的直接链接")
------------------------------------------------------------------------------

此方法用于删除一个 cookie。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `name` | `str` | 必填 | cookie 的 name 字段 |

**返回：**`None`

---

✅️️ `set.headers()`[​](#️️-setheaders "️️-setheaders的直接链接")
-----------------------------------------------------------

此方法用于设置 headers，会取代已有 headers。

headers 可以是`dict`格式的，也可以是文本格式。

文本格式不同字段用`\n`分隔，字段 key 和 value 用`': '`分隔，即从浏览器直接复制的格式。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `headers` | `dict` `str` | 必填 | headers 信息 |

**返回：**`None`

---

✅️️ `set.header()`[​](#️️-setheader "️️-setheader的直接链接")
--------------------------------------------------------

此方法用于设置 headers 中一个项。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `name` | `str` | 必填 | 设置名称 |
| `value` | `str` | 必填 | 设置值 |

**返回：**`None`

---

✅️️ `set.user_agent()`[​](#️️-setuser_agent "️️-setuser_agent的直接链接")
--------------------------------------------------------------------

此方法用于设置 user\_agent。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `ua` | `str` | 必填 | user\_agent 信息 |

**返回：**`None`

---

✅️️ `set.proxies()`[​](#️️-setproxies "️️-setproxies的直接链接")
-----------------------------------------------------------

此方法用于设置代理 ip。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `http` | `str` | `None` | http 代理地址 |
| `https` | `str` | `None` | https 代理地址 |

**返回：**`None`

---

✅️️ `set.auth()`[​](#️️-setauth "️️-setauth的直接链接")
--------------------------------------------------

此方法用于设置认证元组或对象。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `auth` | `Tuple[str, str]` `HTTPBasicAuth` | 必填 | 认证元组或对象 |

**返回：**`None`

---

✅️️ `set.hooks()`[​](#️️-sethooks "️️-sethooks的直接链接")
-----------------------------------------------------

此方法用于设置回调方法。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `hooks` | `dict` | 必填 | 回调方法 |

**返回：**`None`

---

✅️️ `set.params()`[​](#️️-setparams "️️-setparams的直接链接")
--------------------------------------------------------

此方法用于设置查询参数字典。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `params` | `dict` | 必填 | 查询参数字典 |

**返回：**`None`

---

✅️️ `set.verify()`[​](#️️-setverify "️️-setverify的直接链接")
--------------------------------------------------------

此方法用于设置是否验证SSL证书。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_off` | `bool` | 必填 | `bool`表示开或关 |

**返回：**`None`

---

✅️️ `set.cert()`[​](#️️-setcert "️️-setcert的直接链接")
--------------------------------------------------

此方法用于设置SSL客户端证书。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `cert` | `str` `Tuple[str, str]` | 必填 | SSL客户端证书文件的路径(.pem格式)，或(‘cert’, ‘key’)元组 |

**返回：**`None`

---

✅️️ `set.stream()`[​](#️️-setstream "️️-setstream的直接链接")
--------------------------------------------------------

此方法用于设置是否使用流式响应内容。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_off` | `bool` | 必填 | `bool`表示开或关 |

**返回：**`None`

---

✅️️ `set.trust_env()`[​](#️️-settrust_env "️️-settrust_env的直接链接")
-----------------------------------------------------------------

此方法用于设置是否信任环境。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `on_off` | `bool` | 必填 | `bool`表示开或关 |

**返回：**`None`

---

✅️️ `set.max_redirects()`[​](#️️-setmax_redirects "️️-setmax_redirects的直接链接")
-----------------------------------------------------------------------------

此方法用于设置连接最大重定向次数。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| ``times | `int` | 必填 | 最大重定向次数 |

**返回：**`None`

---

✅️️ `set.add_adapter()`[​](#️️-setadd_adapter "️️-setadd_adapter的直接链接")
-----------------------------------------------------------------------

此方法用于添加适配器。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `url` | `str` | 必填 | 适配器对应url |
| `adapter` | `HTTPAdapter` | 必填 | 适配器对象 |

**返回：**`None`

---

✅️️ `close()`[​](#️️-close "️️-close的直接链接")
-------------------------------------------

此方法用于关闭连接。

**参数：** 无

**返回：**`None`

---

[上一页

🛩️ 获取元素信息](/SessionPage/get_ele_info)[下一页

🛩️ 启动配置](/SessionPage/session_opt)

* [✅️️ `set.retry_times()`](#️️-setretry_times)
* [✅️️ `set.retry_interval()`](#️️-setretry_interval)
* [✅️️ `set.timeout()`](#️️-settimeout)
* [✅️️ `set.encoding()`](#️️-setencoding)
* [✅️️ `set.cookies()`](#️️-setcookies)
* [✅  ️️ `set.cookies.clear()`](#️️-setcookiesclear)
* [✅️️ `set.cookies.remove()`](#️️-setcookiesremove)
* [✅️️ `set.headers()`](#️️-setheaders)
* [✅️️ `set.header()`](#️️-setheader)
* [✅️️ `set.user_agent()`](#️️-setuser_agent)
* [✅️️ `set.proxies()`](#️️-setproxies)
* [✅️️ `set.auth()`](#️️-setauth)
* [✅️️ `set.hooks()`](#️️-sethooks)
* [✅️️ `set.params()`](#️️-setparams)
* [✅️️ `set.verify()`](#️️-setverify)
* [✅️️ `set.cert()`](#️️-setcert)
* [✅️️ `set.stream()`](#️️-setstream)
* [✅️️ `set.trust_env()`](#️️-settrust_env)
* [✅️️ `set.max_redirects()`](#️️-setmax_redirects)
* [✅️️ `set.add_adapter()`](#️️-setadd_adapter)
* [✅️️ `close()`](#️️-close)

* 🛫 SessionPage
* 🛩️ 访问网页

本页总览

🛩️ 访问网页
=======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

`SessionPage`基于 requests 进行网络连接，因此可使用 requests 内置的所有请求方式，包括`get()`、`post()`、`head()`、`options()`、`put()`、`patch()`、`delete()`。

不过本库目前只对`get()`和`post()`做了封装和优化，其余方式可通过调用页面对象内置的`Session`对象使用。

✅️️ `get()`[​](#️️-get "️️-get的直接链接")
-------------------------------------

此方法用于以 GET 方式请求页面。

### 📌 访问在线网页[​](#-访问在线网页 "📌 访问在线网页的直接链接")

`get()`方法语法与 requests 的`get()`方法一致，在此基础上增加了连接失败重试功能。

与 requests 不一样的是，它不返回`Response`对象，而是从`SessionPae`对象的`html`等属性读取结果。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `url` | `str` | 必填 | 目标 url，可指向本地文件路径 |
| `show_errmsg` | `bool` | `False` | 连接出错时是否显示和抛出异常 |
| `retry` | `int` | `None` | 重试次数，为`None`时使用页面参数，默认`3` |
| `interval` | `float` | `None` | 重试间隔（秒），为`None`时使用页面参数，默认`2` |
| `timeout` | `float` | `None` | 加载超时时间（秒） |
| `params` | `dict` | `None` | url 请求参数 |
| `data` | `dict` `str` | `None` | 携带的数据 |
| `json` | `dict` `str` | `None` | 要发送的 JSON 数据，会自动设置 Content-Type 为`'application/json'` |
| `headers` | `dict` | `None` | 请求头 |
| `cookies` | `dict` `CookieJar` | `None` | cookies 信息 |
| `files` | `Any` | `None` | 要上传的文件，可以是一个字典，其中键是文件名，值是文件对象或文件路径 |
| `auth` | `Any` | `None` | 身份认证信息 |
| `allow_redirects` | `bool` | `True` | 是否允许重定向 |
| `proxies` | `dict` | `None` | 代理信息 |
| `hooks` | `Any` | `None` | 回调方法 |
| `stream` | `bool` | `None` | 是否使用流式传输 |
| `verify` | `bool` `str` | `None` | 是否验证 SSL 证书 |
| `cert` | `str` `Tuple[str, str]` | `None` | SSL 客户端证书文件的路径(.pem 格式)，或('cert', 'key')元组 |

| 返回类型 | 说明 |
| --- | --- |
| `bool` | 是否连接成功 |

说明

`**kwargs`参数与 requests 中该参数使用方法一致，但有一个特点，如果该参数中设置了某一项（如`headers`），该项中的每个项会覆盖从配置中读取的同名项，而不会整个覆盖。  
就是说，如果想继续使用配置中的`headers`信息，而只想修改其中一项，只需要传入该项的值即可。这样可以简化代码逻辑。

程序会根据要访问的网址自动在`headers`中加入`Host`和`Referer`项

程序会自动从返回内容中确定编码，一般情况无需手动设置

普通访问网页：

```
from DrissionPage import SessionPage  
  
page = SessionPage()  
page.get('http://g1879.gitee.io/drissionpage')
```

使用连接参数访问网页：

```
from DrissionPage import SessionPage  
  
page = SessionPage()  
  
url = 'https://www.baidu.com'  
headers = {'referer': 'gitee.com'}  
cookies = {'name': 'value'}  
proxies = {'http': '127.0.0.1:1080', 'https': '127.0.0.1:1080'}  
page.get(url, headers=headers, cookies=cookies, proxies=proxies)
```

---

### 📌 读取本地文件[​](#-读取本地文件 "📌 读取本地文件的直接链接")

`get()`的`url`参数可指向本地文件，实现本地 html 解析。

```
from DrissionPage import SessionPage  
  
page = SessionPage()  
page.get(r'D:\demo.html')
```

---

✅️️ `post()`[​](#️️-post "️️-post的直接链接")
----------------------------------------

此方法是用 POST 方式请求页面。用法与`get()`一致。

| 参数名称 | 类型 | 默认值 | 说明 |
| --- | --- | --- | --- |
| `url` | `str` | 必填 | 目标 url，可指向本地文件路径 |
| `show_errmsg` | `bool` | `False` | 连接出错时是否显示和抛出异常 |
| `retry` | `int` | `None` | 重试次数，为`None`时使用页面参数，默认`3` |
| `interval` | `float` | `None` | 重试间隔（秒），为`None`时使用页面参数，默认`2` |
| `timeout` | `float` | `None` | 加载超时时间（秒） |
| `params` | `dict` | `None` | url 请求参数 |
| `data` | `dict` `str` | `None` | 携带的数据 |
| `json` | `dict` `str` | `None` | 要发送的 JSON 数据，会自动设置 Content-Type 为`'application/json'` |
| `headers` | `dict` | `None` | 请求头 |
| `cookies` | `dict` `CookieJar` | `None` | cookies 信息 |
| `files` | `Any` | `None` | 要上传的文件，可以是一个字典，其中键是文件名，值是文件对象或文件路径 |
| `auth` | `Any` | `None` | 身份认证信息 |
| `allow_redirects` | `bool` | `True` | 是否允许重定向 |
| `proxies` | `dict` | `None` | 代理信息 |
| `hooks` | `Any` | `None` | 回调方法 |
| `stream` | `bool` | `None` | 是否使用流式传输 |
| `verify` | `bool` `str` | `None` | 是否验证 SSL 证书 |
| `cert` | `str` `Tuple[str, str]` | `None` | SSL 客户端证书文件的路径(.pem 格式)，或('cert', 'key')元组 |

| 返回类型 | 说明 |
| --- | --- |
| `bool` | 是否连接成功 |

```
from DrissionPage import SessionPage  
  
page = SessionPage()  
data = {'username': '****', 'pwd': '****'}  
  
page.post('http://example.com', data=data)  
# 或  
page.post('http://example.com', json=data)
```

`data`参数和`json`参数都可接收`str`和`dict`格式数据，即有以下 4 种传递数据的方式：

```
# 向 data 参数传入字符串  
page.post(url, data='abc=123')  
  
# 向 data 参数传入字典  
page.post(url, data={'abc': '123'})  
  
# 向 json 参数传入字符串  
page.post(url, json='abc=123')  
  
# 向 json 参数传入字典  
page.post(url, json={'abc': '123'})
```

具体使用哪种，按服务器要求而定。

---

✅️️ 其它请求方式[​](#️️-其它请求方式 "✅️️ 其它请求方式的直接链接")
-------------------------------------------

本库只针对常用的 GET 和 POST 方式作了优化，但也可以通过提取页面对象内的`Session`对象以原生 requests 代码方式执行其它请求方式。

```
from DrissionPage import SessionPage  
  
page = SessionPage()  
# 获取内置的 Session 对象  
session = page.session  
# 以 head 方式发送请求  
response = session.head('https://www.baidu.com')  
print(response.headers)
```

**输出：**

```
{'Accept-Ranges': 'bytes', 'Cache-Control': 'private, no-cache, no-store, proxy-revalidate, no-transform', 'Connection': 'keep-alive', 'Content-Length': '277', 'Content-Type': 'text/html', 'Date': 'Tue, 04 Jan 2022 06:49:18 GMT', 'Etag': '"575e1f72-115"', 'Last-Modified': 'Mon, 13 Jun 2016 02:50:26 GMT', 'Pragma': 'no-cache', 'Server': 'bfe/1.0.8.18'}
```

[上一页

🛩️ 创建页面对象](/SessionPage/create_obj)[下一页

🛩️ 获取页面信息](/SessionPage/get_page_info)

* [✅️️ `get()`](#️️-get)
  + [📌 访问在线网页](#-访问在线网页)
  + [📌 读取本地文件](#-读取本地文件)
* [✅️️ `post()`](#️️-post)
* [✅️️ 其它请求方式](#️️-其它请求方式)

## <a name="教程"></a>教程

* ❓ 常见问题

本页总览

❓ 常见问题
======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

本页收集一些用户使用过程中的常见问题。

欢迎各位开发者添砖加瓦，您可提交 issues、PR，也可以写成博客文章，链接发给本库作者，直接链接到文章。

❔ 如何在无界面 Linux 使用[​](#-如何在无界面-linux-使用 "❔ 如 何在无界面 Linux 使用的直接链接")
-----------------------------------------------------------------

CentOS 请参考这篇文章：[linux 部署说明](https://blog.csdn.net/sinat_39327967/article/details/132181129?spm=1001.2014.3001.5501)

Ubuntu 请参考这篇文章：[DrissionPage在UbuntuLinux的使用](https://zhuanlan.zhihu.com/p/674687748)

---

❔ 为什么浏览器不能退出无头模式？[​](#-为什么浏览器不能退出无头模式 "❔ 为什么浏览器不能退出无头模式？的直接链接")
---------------------------------------------------------------

> 为什么设置过无头后，下次运行即使不设置`headless()`，浏览器依然进入无头状态？

因为上一次打开的浏览器没有关闭，只是因为开了无头不可见，程序继续接管了它。

如果要关闭浏览器，可在程序结束时使用`Chromium`对象的`quit()`语句。

也可以设置`co.headless(False)`，程序会自动关闭之前的无头浏览器再启动新的。

另请注意，`tab.close()`的功能是关闭当前标签页，而不是关闭浏览器，除非浏览器只有一个标签页。

---

❔ 如何禁用保存密码、恢复页面等提示气泡？[​](#-如何禁用保存密码恢复页面等提示气泡 "❔ 如何禁用保存密码、恢复页面等提示气泡？的直接链接")
--------------------------------------------------------------------------

浏览器提示气泡出现时可以手动关闭，不关闭也不影响自动操作，在代码中阻止其显示也是可以的。
加一些浏览器配置代码即可禁止相应的气泡显示，需要添加下面这样的代码：

```
co = ChromiumOptions()  
  
# 阻止“自动保存密码”的提示气泡  
co.set_pref('credentials_enable_service', False)  
  
# 阻止“要恢复页面吗？Chrome未正确关闭”的提示气泡  
co.set_argument('--hide-crash-restore-bubble')
```

---

❔ 点击报错“该元素没有位置及大小”怎么办？[​](#-点击报错该元素没有位置及大小怎么办 "❔ 点击报错“该元素没有位置及大小”怎么办？的直接链接")
----------------------------------------------------------------------------

没有位置及大小是正常的，很多元素都没有位置和大小。

这个时候你要检查是否页面中有同名元素，定位符没写准确拿到了另一个。

如果要点击的元素就是没有位置的，可以强制使用 js 点击，用法是 `.click(by_js=True)`，可以简写为 `.click('js')`。

---

❔ 如何使用启动参数、用户配置、实验项等功能？[​](#-如何使用启动参数用户配置实验项等功能 "❔ 如何使用启动参数、用户配置、实验项等功能？的直接链接")
-------------------------------------------------------------------------------

**arguments 启动参数**

* 使用参考：<http://DrissionPage.cn/ChromiumPage/browser_opt#-set_argument>
* 参数详见：<https://peter.sh/experiments/chromium-command-line-switches/>

**prefs 用户配置**

* 使用参考：<http://DrissionPage.cn/ChromiumPage/browser_opt#-set_pref>
* 参数详见：<https://src.chromium.org/viewvc/chrome/trunk/src/chrome/common/pref_names.cc>

**flags 实验项**

* 使用参考：<http://DrissionPage.cn/ChromiumPage/browser_opt#-set_flag>
* 参数详见：chrome://flags

注意

外部链接仅供参考，请谨慎使用任何高级功能，仅在确保一切都可以掌控时才可使用，因为使用这些功能可能会导致浏览器数据丢失或安全和隐私受到威胁。

---

❔ 如何匹配特殊字符（如`'&nbsp;'`）文本？[​](#-如何匹配特殊字符如nbsp文本 "-如何匹配特殊字符如nbsp文本的直接链接")
------------------------------------------------------------------------

需先将特殊字符转为十六进制形式，详见《查找元素》中《语法速查表》一节。

[上一页

🗨️ 公众号](/tutorials/gongzhonghao)[下一页

🥦 创建全新的浏览器](/tutorials/functions/new_browser)

* [❔ 如何在无界面 Linux 使用](#-如何在无界面-linux-使用)
* [❔ 为什么浏览器不能退出无头模式？](#-为什么浏览器不能退出无头模式)
* [❔ 如何禁用保存密码、恢复页 面等提示气泡？](#-如何禁用保存密码恢复页面等提示气泡)
* [❔ 点击报错“该元素没有位置及大小”怎么办？](#-点击报错该元素没有位置及大小怎么办)
* [❔ 如何使用启动参数、用户配置、实验项等功能？](#-如何使用启动参数用户配置实验项等功能)
* [❔ 如何匹配特殊字符（如`'&nbsp;'`）文本？](#-如何匹配特殊字符如nbsp文本)

本页总览

🏪 小店
====

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

欢迎光顾作者的 B 站小店[​](#欢迎光顾作者的-b-站小店 "欢迎光顾作者的 B 站小店的直接链接")
-----------------------------------------------------

作者的 B 站小店会陆续上线一些实用工具、离线文档、案例代码等，欢迎光临。

![](/assets/images/barrack-d92b34e66fbfb806f484614a6b8f8833.png)

---

商品[​](#商品 "商品的直接链接")
--------------------

[骚神插件](https://gf.bilibili.com/item/detail/1108443001)

[DrissionPage 离线文档](https://gf.bilibili.com/item/detail/1108367001)

* [欢迎光顾作者的 B 站小店](#欢迎光顾作者的-b-站小店)
* [商品](#商品)

* 🥬 功能示例
* 🥦 浏览器多开

本页总览

🥦 浏览器多开
=======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

本节介绍如何启动多个浏览器同时使用。

✅️️ 直接指定端口[​](#️️-直接指定端口 "✅️️ 直接指定端口的直接链接")
-------------------------------------------

使用`Chromium()`启动浏览器时可以直接指定端口来启动多个浏览器。

如果指定端口已有浏览器在运行，会接管这个浏览器。

使用端口号创建的浏览器用户数据文件夹会保留，只要临时文件夹未被清理，下次使用该端口时还会使用这些数据，比如登录信息和插件。

```
from DrissionPage import Chromium  
  
browser1 = Chromium(9222)   
browser2 = Chromium(9333)
```

✅️️ 自动设置端口[​](#️️-自动设置端口 "✅️️ 自动设置端口的直接链接")
-------------------------------------------

使用`ChromiumOptions`对象的`auto_port()`方法，可自动获取空闲端口，并创建全新浏览器（无用户数据和插件）。

这时多个页面对象可共用一个`ChromiumOptions`对象，不会产生冲突。

浏览器关闭后会自动删除用户文件夹，不会过多占用硬盘空间。

```
from DrissionPage import Chromium, ChromiumOptions  
  
co = ChromiumOptions().auto_port()  
browser1 = Chromium(co)  
browser2 = Chromium(co)
```

✅️️ 手动指定端口和路径[​](#️️-手动指定端口和路径 "✅️️ 手动指定端口和路径的直接链接")
----------------------------------------------------

也可以用指定独立端口和用户文件夹路径的方式启动多个浏览器。

需要注意的是，端口和用户文件夹每个浏览器都要独立使用，不能共用。

因此，这时每个页面对象需要自己的配置对象。

```
from DrissionPage import Chromium, ChromiumOptions  
  
co1 = ChromiumOptions().set_local_port(9222).set_user_data_path('data1')  
co2 = ChromiumOptions().set_local_port(9333).set_user_data_path('data2')  
browser1 = Chromium(co1)  
browser2 = Chromium(co2)
```

[上一页

🥦 创建全新的浏览器](/tutorials/functions/new_browser)[下一页

🥦 无头模式](/tutorials/functions/headless)

* [✅️️ 直接指定端口](#️️-直接指定端口)
* [✅️️ 自动设置端口](#️️-自动设置端口)
* [✅️️ 手动指定端口和路径](#️️-手动指定端口和路径)

* 🥬 功能示例
* 🥦 无头模式

🥦 无头模式
======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

要使用无头模式很简单，在`ChromiumOptions`设置`headless()`即可。

```
from DrissionPage import Chromium, ChromiumOptions  
  
co = ChromiumOptions().headless()  
browser = Chromium(co)
```

需要注意的是，程序结束时浏览器不会自动关闭，下次运行会继续接管该浏览器。

无头浏览器因为看不见很容易被忽视。可在程序结尾用`browser.quit()`将其关闭。

[上一页

🥦 浏览器多开](/tutorials/functions/create_browsers)[下一页

🥦 设置 cookies](/tutorials/functions/set_cookies)

* 🥬 功能示例
* 🥦 创建全新的浏览器

本页总览

🥦 创建全新的浏览器
==========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

默认设置下，DrissionPage 会在 9222 端口创建浏览器，如果该端口下浏览器已经启动，则会接管使用。

并且会重复使用用户文件夹，即该端口登录过的账号，下次启动时可能仍有效，使用的插件也一样。

这为调试和日常使用带来大量便利，无须总是要处理登录和加载插件。

但项目中往往需要创建全新的浏览器环境，不希望复用之前的用户数据，可用以下方法实现：

✅️️ `auto_port()`方法[​](#️️-auto_port方法 "️️-auto_port方法的直接链接")
-------------------------------------------------------------

使用`ChromiumOptions`对象的`atuo_port()`方法，可指定程序自动创建全新的浏览器，多个浏览器互不干扰。

```
from DrissionPage import ChromiumOptions, Chromium  
  
co = ChromiumOptions().auto_port()  
browser1 = Chromium(co)  
browser2 = Chromium(co)
```

如此即可创建两个全新且独立的浏览器。

可以注意到，示例中两个`Chromium`对象共用了一个`ChromiumOptions`对象，这在设置`auto_port()`时才会生效。

如果没有设置`auto_port()`，两个页面对象其实是同一个。

注意

如果使用`atuo_port()`后再使用`set_local_port()`、`set_address()`或`set_user_data_path()`，会覆盖`auto_port()`设置。

---

✅️️ 手动配置[​](#️️-手动配置 "✅️️ 手动配置的直接链接")
-------------------------------------

如果有更细致的需求，不使用`auto_port()`，可自行使用`set_local_port()`、`set_address()`和`set_user_data_path()`为每个浏览器指定端口和用户文件夹。

注意

* 务必注意的是，每个浏览器的端口和用户文件夹都必须是独立的，不能复用
* 每个浏览器都要一个`ChromiumOptions`对象，不能复用

```
from DrissionPage import Chromium, ChromiumOptions  
  
co1 = ChromiumOptions().set_local_port(9111).set_user_data_path('data1')  
co2 = ChromiumOptions().set_local_port(9222).set_user_data_path('data2')  
  
browser1 = Chromium(co1)  
browser2 = Chromium(co2)
```

[上一页

❓ 常见问题](/tutorials/QandA)[下一页

🥦 浏览器多开](/tutorials/functions/create_browsers)

* [✅️️ `auto_port()`方法](#️️-auto_port方法)
* [✅️️ 手动配置](#️️-手动配置)

* 🥬 功能示例
* 🥦 设置 cookies

本页总览

🥦 设置 cookies
============

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

✅️️ 设置 cookies[​](#️️-设置-cookies "✅️️ 设置 cookies的直接链接")
-------------------------------------------------------

### 📌 页面对象中设置[​](#-页面对象中设置 "📌 页面对象中设置的直接链接")

任意页面对象都有`set.cookies()`方法，用于设置 cookies。

该方法接收多种格式的 cookies 信息，可设置一个或多个 cookies。

使用浏览器时，任意页面对象设置的 cookies 是所有标签页共用的（由`new_tab(new_context=True)`创建的标签页除外）。

**示例：**

```
from DrissionPage import Chromium  
  
tab = Chromium().latest_tab  
cookies = 'name1=value1; name2=value2; path=/; domain=.example.com;'  
  
tab.set.cookies(cookies)
```

---

### 📌 `SessionOptions`中设置[​](#-sessionoptions中设置 "-sessionoptions中设置的直接链接")

`SessionOptions`对象有`set_cookies()`方法，可接收一个或多个 cookies，用于`SessionPage`初始化时设置 cookies。

每次设置会覆盖之前所有 cookies 信息。

**示例：**

```
from DrissionPage import SessionOptions  
  
cookies = 'name1=value1; name2=value2; path=/; domain=.example.com;'  
  
co = SessionOptions()  
co.set_cookies(cookies)
```

---

### 📌 删除 cookies[​](#-删除-cookies "📌 删除 cookies的直接链接")

页面对象用`set.cookies.remove()`和`set.cookies.clear()`删除和清空 cookies。

`SessionOptions`对象用`set_cookies(None)`清空 cookies。

具体用法详见使用文档有关章节。

---

✅️️ cookies 格式[​](#️️-cookies-格式 "✅️️ cookies 格式的直接链接")
-------------------------------------------------------

### 📌 设置一个 cookie[​](#-设置一个-cookie "📌 设置一个 cookie的直接链接")

设置一个 cookie 时，可传入`Cookie`、`dict`或`str`类型。

`dict`和`str`需要有`name`和`value`字段。

`str`多个字段间用`';'`或`','`分隔，但不能两种同时出现。

**格式：**

```
# dict类型  
{'name': 'abc', 'value': '123', 'domain': '.example.com', ...}  
  
# str类型  
'name=abc; value=123; domain=.example.com; ...'
```

---

### 📌 设置多个 cookies[​](#-设置多个-cookies "📌 设置多个 cookies的直接链接")

设置多个时，可传入`CookieJar`、`list`、`tuple`、`str`、`dict`类型。

列表里面可以放`Cookie`、`str`或`dict`类型，多个 cookies 格式可以是不同的。

注意

列表中如果放`str`或`dict`，每个项都只能是一个 cookie。

**格式：**

```
# dict类型  
{'abc': '123', 'def': '456', 'domain': '.example.com', ...}  
  
# str类型  
'abc=123; def=456; domain=.example.com; ...'  
  
# list或tuple类型  
['name=123; domain=.example.com; ...', 'name=abc; value=123; domain=.example.com; ...']
```

---

### 📌 说明[​](#-说明 "📌 说明的直接链接")

cookies 中只有`name`和`value`字段是必须的，但如果没有`domain`字段，添加到浏览器时会自动添加。

添加的内容根据调用`set.cookies()`方法的对象 url 而定。

比如一个 Tab 对象当前 url 为`'https://www.baidu.com'`，添加无指定域名的 cookies 时，会自动添加该字段，内容为`'www.baidu.com'`。

[上一页

🥦 无头模式](/tutorials/functions/headless)

* [✅️️ 设置 cookies](#️️-设置-cookies)
  + [📌 页面对象中设置](#-页面对象中设置)
  + [📌 `SessionOptions`中设置](#-sessionoptions中设置)
  + [📌 删除 cookies](#-删除-cookies)
* [✅️️ cookies 格式](#️️-cookies-格式)
  + [📌 设置一个 cookie](#-设置一个-cookie)
  + [📌 设置多个 cookies](#-设置多个-cookies)
  + [📌 说明](#-说明)

* 🗨️ 公众号

本页总览

🗨️ 公众号
======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

欢迎关注公众号[​](#欢迎关注公众号 "欢迎关注公众号的直接链接")
-----------------------------------

![](/assets/images/gzh-2ebd3a4f84055229e0db2cfbaf02558f.jpg)

[上一页

🎞️ 视频教程](/tutorials/video)[下一页

❓ 常见问题](/tutorials/QandA)

* [欢迎关注公众号](#欢迎关注公众号)

* 🎞️ 视频教程

🎞️ 视频教程
=======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

* Bilibili：<https://space.bilibili.com/20526000>
* 小红书：<https://www.xiaohongshu.com/user/profile/63293b09000000002303e82e>
* 抖音号：52635622816

![](/assets/images/video-318e4be75d5fe6df2af02901a317f02d.png)

[上一页

🌏️ 知识星球](/tutorials/xingqiu)[下一页

🗨️ 公众号](/tutorials/gongzhonghao)

* 🌏️ 知识星球

本页总览

🌏️ 知识星球
=======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

欢迎加入作者的知识星球[​](#欢迎加入作者的知识星球 "欢迎加入作者的知识星球的直接链接")
-----------------------------------------------

这里是一个关于 DrissionPage 实战案例、问题记录、解决问题、共同讨论的地方。

在这里，你可以：

* 找到 dp 常见的问题的解决方案
* 学到更多使用技巧和案例
* 向作者提问，获得官方技术支持
* 查看作者的开发笔记
* 与志同道合的其他人一起讨论技术，共同成长

还欢迎提出好的点子和改进意见，把 dp 打造得越来越好。

![](/assets/images/zsxq-d2326c68adc888e01dda6b8a75721ec7.jpg)

[下一页

🎞️ 视频教程](/tutorials/video)

* [欢迎加入作者的知识星球](#欢迎加入作者的知识星球)

## <a name="版本"></a>版本

* 📒 v0.x-v1.4

本页总览

📒 v0.x-v1.4
===========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

v0.x 至 v1.4 版本是基于 selenium 和 requests-html 制作，前者负责控制浏览器部分，后者负责收发数据包部分。

v1.4.0[​](#v140 "v1.4.0的直接链接")
------------------------------

* d 模式使用 js 通过`evaluate()`方法处理 xpath，放弃使用 selenium 原生的方法，以支持用 xpath 直接获取文本节点、元素属性
* d 模式增加支持用 xpath 获取元素文本、属性
* 优化和修复小问题

v1.3.0[​](#v130 "v1.3.0的直接链接")
------------------------------

* 可与 selenium 代码无缝对接
* 下载功能支持 post 方式
* 元素添加`texts`属性，返回元素内每个文本节点内容
* s 模式增加支持用 xpath 获取元素文本、属性

v1.2.1[​](#v121 "v1.2.1的直接链接")
------------------------------

* 优化修复网页编码逻辑
* `download()`函数优化获取文件名逻辑
* 优化`download()`获取文件大小逻辑
* 优化`MixPage`对象关闭 session 逻辑

v1.2.0[​](#v120 "v1.2.0的直接链接")
------------------------------

* 增加对 shadow-root 的支持
* 增加自动重试连接功能
* `MixPage`可直接接收配置
* 修复一些 bug

v1.1.3[​](#v113 "v1.1.3的直接链接")
------------------------------

* 连接有关函数增加是否抛出异常参数
* s 模式判断编码优化
* d 模式`check_page()`优化
* 修复`run_script()`遗漏`args`参数的问题

v1.1.1[​](#v111 "v1.1.1的直接链接")
------------------------------

* 删除`get_tabs_sum()`和`get_tab_num()`函数，以属性`tabs_count`和`current_tab_num`代替
* 增加`current_tab_handle`、`tab_handles`属性
* `to_tab()`和`close_other_tabs()`函数可接收`handle`值
* `create_tab()`可接收一个 url 在新标签页打开
* 其它优化和 bug 修复

v1.1.0[​](#v110 "v1.1.0的直接链接")
------------------------------

* 元素对象增加 xpath 和 css path 路径属性
* 修复 driver 模式下元素对象用 css 方式不能获取直接子元素的问题（selenium 的锅）
* s 模式下现在能通过 xpath 定位上级元素
* 优化 d 模式兄弟元素、父级元素的获取效率
* 优化标签页处理功能
* 其它小优化和修复

v1.0.5[​](#v105 "v1.0.5的直接链接")
------------------------------

* 修复切换模式时 url 出错的 bug

v1.0.3[​](#v103 "v1.0.3的直接链接")
------------------------------

* `DriverOptions`支持链式操作
* `download()`函数增加参数处理遇到已存在同名文件的情况，可选跳过、覆盖、自动重命名
* `download()`函数重命名调整为只需输入文件名，不带后缀名，输入带后缀名也可自动识别

v1.0.1[​](#v101 "v1.0.1的直接链接")
------------------------------

* 增强拖拽功能和 chrome 设置功能

v0.14.0[​](#v0140 "v0.14.0的直接链接")
---------------------------------

* `Drission`类增加代理设置和修改

v0.12.4[​](#v0124 "v0.12.4的直接链接")
---------------------------------

* `click()`的`by_js`可接收`False`
* 修复一些 bug

v0.12.0[​](#v0120 "v0.12.0的直接链接")
---------------------------------

* 增加`tag:tagName@arg=val`查找元素方式
* `MixPage`增加简易方式创建对象方式

v0.11.0[​](#v0110 "v0.11.0的直接链接")
---------------------------------

* 完善`easy_set`的函数
* 元素增加多级定位函数

v0.10.2[​](#v0102 "v0.10.2的直接链接")
---------------------------------

* 完善`attr`及`attrs`功能

v0.10.1[​](#v0101 "v0.10.1的直接链接")
---------------------------------

* 增加`set_headless()`以及`to_iframe()`兼容全部原生参数

v0.9.4[​](#v094 "v0.9.4的直接链接")
------------------------------

* 修复 bug

v0.9.0[​](#v090 "v0.9.0的直接链接")
------------------------------

* 增加了元素拖拽和处理提示框功能

v0.8.4[​](#v084 "v0.8.4的直接链接")
------------------------------

* 基本完成

[上一页

📒 v1.5-v2.x](/versions/2x)[下一页

📆 下一步计划](/versions/next)

* [v1.4.0](#v140)
* [v1.3.0](#v130)
* [v1.2.1](#v121)
* [v1.2.0](#v120)
* [v1.1.3](#v113)
* [v1.1.1](#v111)
* [v1.1.0](#v110)
* [v1.0.5](#v105)
* [v1.0.3](#v103)
* [v1.0.1](#v101)
* [v0.14.0](#v0140)
* [v0.12.4](#v0124)
* [v0.12.0](#v0120)
* [v0.11.0](#v0110)
* [v0.10.2](#v0102)
* [v0.10.1](#v0101)
* [v0.9.4](#v094)
* [v0.9.0](#v090)
* [v0.8.4](#v084)

* 📒 v1.5-v2.x

本页总览

📒 v1.5-v2.x
===========

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

v1.5 至 v2.x 版本，基于 selenium 控制浏览器，用作者自制的功能收发数据包。

v2.7.3[​](#v273 "v2.7.3的直接链接")
------------------------------

* 页面对象和元素对象的`screenshot_as_bytes()`方法合并到`screenshot()`
* `input()`方法接收非文本参数时自动转成文本输入

v2.7.2[​](#v272 "v2.7.2的直接链接")
------------------------------

* d 模式页面和元素对象增加`screenshot_as_bytes()`方法

v2.7.1[​](#v271 "v2.7.1的直接链接")
------------------------------

* DriverPage
  + 增加`get_session_storage()`、`get_local_storage()`、`set_session_storage()`、`set_local_storage()`、`clean_cache()`方法
  + `run_cdp()`的`cmd_args`参数改为`**cmd_args`
* 关闭 driver 时会主动关闭 chromedriver.exe 的进程
* 优化关闭浏览器进程逻辑

v2.6.2[​](#v262 "v2.6.2的直接链接")
------------------------------

* d 模式增加`stop_loading()`方法
* 优化完善监听器功能

v2.6.0[​](#v260 "v2.6.0的直接链接")
------------------------------

* 新增`Listener`类
  + 可监听浏览器数据包
  + 可异步监听
  + 可实现每监听到若干数据包执行操作
* 放弃对selenium4.1以下的支持
* 解决使用新版浏览器时出现的一些问题

v2.5.9[​](#v259 "v2.5.9的直接链接")
------------------------------

* 优化 s 模式创建连接的逻辑

v2.5.7[​](#v257 "v2.5.7的直接链接")
------------------------------

* 列表元素`select()`、`deselect()`等方法添加`timeout`参数，可等待列表元素加载
* 优化了对消息提示框的处理
* `drag()`和`drag_to()`不再检测是否拖拽成功，改成返回`None`
* `DriverOptions`对象从父类继承的方法也支持链式操作
* 其它优化和问题修复

v2.5.5[​](#v255 "v2.5.5的直接链接")
------------------------------

* `DriverPage`添加`run_cdp()`方法
* `get()`和`post()`方法删除`go_anyway`参数
* 连接重试默认不打印提示

v2.5.0[​](#v250 "v2.5.0的直接链接")
------------------------------

* 用 DownloadKit 库替代原来的`download()`方法，支持多线程并发
* `DriverPage`增加`set_ua_to_tab()`方法
* 删除`scroll_to()`方法
* 其它  优化和问题修复

v2.4.3[​](#v243 "v2.4.3的直接链接")
------------------------------

* `wait_ele()`、`to_frame()`、`scroll_to()`改用类的方式，避免使用字符串方式选择功能
* `scroll_to()`方法改为`scroll`属性
* 滚动页面或元素增加`to_location()`方式
* 优化`Select`类用法

v2.3.3[​](#v233 "v2.3.3的直接链接")
------------------------------

* `DriverPage`添加`forward()`方法
* `DriverPage`的`close_current_tab()`改为`close_tabs()`，可一次过关闭多个标签页
* `DriverPage`添加`run_async_script()`
* `DriverPage`添加`timeouts`属性
* `DriverPage`添加`set_timeouts()`方法
* `DriverElement`添加`scroll_to()`方法，可在元素内滑动滚动条
* `DriverOptions`添加`set_page_load_strategy()`方法
* ini 文件增加`page_load_strategy`、`set_window_rect`、`timeouts`三个属性
* 其它优化和问题修复

v2.2.1[​](#v221 "v2.2.1的直接链接")
------------------------------

* 新增基于页面布局的相对定位方法`left()`，`right()`，`below()`，`above()`，`near()`，`lefts()`，`rights()`，`belows()`，`aboves()`，`nears()`
* 修改基于 DOM 的相对定位方法：删除`parents()`方法，`parent`属性改为 `parent()`方法，`next`属性 改为`next()`方法，`prev`属性改为`prev()`方法，`nexts()`和`prevs()`
  方法改为返回多个对象
* 增加`after()`，`before()`，`afters()`，`before()`等基于 DOM 的相对定位方法
* 定位语法增加`@@`和`@@-`语法，用于同时匹配多个条件和排除条件
* 改进`ShadowRootElement`功能，现在在 shadow-root 下查找元素可用完全版的定位语法。
* `DriverElement`的`after`和`before`属性改为`pseudo_after`和`pseudo_before`
* `DriverElement`的`input()`增加`timeout`参数
* `DriverElement`的`clear()`增加`insure_clear`参数
* 优化`DriverElement`的`submit()`方法
* `DriverPage`增加`active_ele`属性，获取焦点所在元素
* `DriverPage`的`get_style_property()`改名为`style()`
* `DriverPage`的`hover()`增加偏移量参数
* `DriverPage`的`current_tab_num`改名为`current_tab_index`
* `DriverPage`的`to_frame()`方法返回页面对象自己，便于链式操作
* 优化自动下载 driver 逻辑
* `set_paths()`增加`local_port`参数
* 默认使用`9222`端口启动浏览器
* 其它优化和问题修复

v2.0.0[​](#v200 "v2.0.0的直接链接")
------------------------------

* 支持从`DriverElement`或 html 文本生成`SessionElement`，可把 d 模式的页面信息爬取速度提高几个数量级（使用新增的`s_ele()`和`s_eles()`方法）
* 支持随时隐藏和显示浏览器进程窗口（只支持 Windows 系统）
* s 模式和 d 模式使用相同的提取文本逻辑，d 模式提取文本效率大增
* `input()`能自动检测以确保输入成功
* `click()`支持失败后不断重试，可用于确保点击成功及等待页面遮罩层消失
* 对 linux 和 mac 系统路径问题做了修复
* `download()`能更准确地获取文件名
* 其它稳定性和效率上的优化

v1.11.7[​](#v1117 "v1.11.7的直接链接")
---------------------------------

* `SessionOptions`增加`set_headers()`方法
* 调整`MixPage`初始化参数
* `click()`增加`timeout`参数，修改逻辑为在超时时间内不断重试点击。可用于监视遮罩层是否消失
* 处理`process_alert()`增加`timeout`参数
* 其他优化和问题修复

v1.11.0[​](#v1110 "v1.11.0的直接链接")
---------------------------------

* `set_property()`方法改名为`set_prop`
* 增加`prop()`
* `clear()`改用 selenium 原生
* 增加`r_click()`和`r_click_at()`
* `input()`返回`None`
* 增加`input_txt()`

v1.10.0[​](#v1100 "v1.10.0的直接链接")
---------------------------------

* 优化启动浏览器的逻辑
* 用 debug 模式启动时可读取启动参数
* 完善`select`标签处理功能
* `MixPage`类的`to_iframe()`改名为`to_frame()`
* `MixPage`类的`scroll_to()`增加`'half'`方式，滚动半页
* Drission 类增加`kill_browser()`方法

v1.9.0[​](#v190 "v1.9.0的直接链接")
------------------------------

* 元素增加`click_at()`方法，支持点击偏移量
* `download()`支持重试
* 元素`input()`允许接收组合键，如`ctrl+a`
* 其它优化

v1.8.0[​](#v180 "v1.8.0的直接链接")
------------------------------

* 添加`retry_times`和`retry_interval`属性，可统一指定重连次数
* 元素对象增加`raw_text`属性
* 元素查找字符串支持极简模式，用`x`表示`xpath`、`c`表示`css`、`t`表示`tag`、`tx`表示`text`
* s 模式元素`text`尽量与 d 模式保持一致
* 其它完善和问题修复

v1.7.7[​](#v177 "v1.7.7的直接链接")
------------------------------

* 创建`WebDriver`时可自动下载 chromedriver.exe
* 修复获取不到`content-type`时会出现的问题

v1.7.1[​](#v171 "v1.7.1的直接链接")
------------------------------

* d 模式如指定了调试端口，可自动启动浏览器进程并接入
* 去除对 cssselect 库依赖
* 提高查找元素效率
*  调整获取元素 xpath 和 css\_path 逻辑

v1.7.0[​](#v170 "v1.7.0的直接链接")
------------------------------

* 优化`cookies`相关逻辑
* `MixPage`增加`get_cookies()`和`set_cookies()`方法
* 增加`SessionOptions`类
* 浏览文件`DriverElement`增加`remove_attr()`方法
* 修复`MixPage`初始化时`Session`导入`cookies`时的问题
* `MixPage`的`close_other_tabs()`方法现在可以接收列表或元组以保留多个 tab
* 其它优化

v1.6.1[​](#v161 "v1.6.1的直接链接")
------------------------------

* 增加`.`和`#`方式用于查找元素，相当于`@Class`和`@id`
* easy\_set 增加识别 chrome 版本并自动下载匹配的 chromedriver.exe 功能
* 改进配置功能
* 修复 shadow-root 方面的问题

v1.5.4[​](#v154 "v1.5.4的直接链接")
------------------------------

* 优化获取编码的逻辑
* 修复下载不能显示进度的问题

v1.5.2[​](#v152 "v1.5.2的直接链接")
------------------------------

* 修复获取 html 时会把元素后面的文本节点带上的问题
* 修复获取编码可能出现的错误
* 优化`download()`和获取编码代码

v1.5.1[​](#v151 "v1.5.1的直接链接")
------------------------------

* 修复获取编码可能出现的 bug

v1.5.0[​](#v150 "v1.5.0的直接链接")
------------------------------

* s 模式使用 lxml 库代替 requests\_html 库
* 可直接调用页面对象和元素对象获取下级元素，`element('@id=ele_id')`等价于`element.ele('@id=ele_id')`
* `nexts()`、`prevs()`方法可获取文本节点
* 可获取伪元素属性及文本
* 元素对象增加`link`和`inner_html`属性
* 各种优化

[上一页

📒 v3.x](/versions/3x)[下一页

📒 v0.x-v1.4](/versions/1x)

* [v2.7.3](#v273)
* [v2.7.2](#v272)
* [v2.7.1](#v271)
* [v2.6.2](#v262)
* [v2.6.0](#v260)
* [v2.5.9](#v259)
* [v2.5.7](#v257)
* [v2.5.5](#v255)
* [v2.5.0](#v250)
* [v2.4.3](#v243)
* [v2.3.3](#v233)
* [v2.2.1](#v221)
* [v2.0.0](#v200)
* [v1.11.7](#v1117)
* [v1.11.0](#v1110)
* [v1.10.0](#v1100)
* [v1.9.0](#v190)
* [v1.8.0](#v180)
* [v1.7.7](#v177)
* [v1.7.1](#v171)
* [v1.7.0](#v170)
* [v1.6.1](#v161)
* [v1.5.4](#v154)
* [v1.5.2](#v152)
* [v1.5.1](#v151)
* [v1.5.0](#v150)

* 📒 v3.x

本页总览

📒 v3.x
======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

v3.2.35[​](#v3235 "v3.2.35的直接链接")
---------------------------------

* 修复浏览器窗口最小化时不响应模拟操作的问题
* 接管浏览器无需`'--remote-allow-origins=*'`参数
* `tabs`属性忽略隐私声明
* 修复 8x 版浏览器选择下拉列表时报错问题
* 修复某些情况下下拉框不触发联动的问题
* 修复配置文件损坏时出现的问题
* 修复`get()`方法`url`参数含某些特殊字符时连接失败的问题

v3.2.33[​](#v3233 "v3.2.33的直接链接")
---------------------------------

* 无界面 Linux 自动进入无头模式
* 添加 MAC 和 Linux 系统默认浏览器路径
* 修复元素截图时可能出现的问题
* 修复`quit()`没有正确等待浏览器进程结束问题
* 屏蔽 MAC 和 Linux 系统多余的提示
* 修复`set.timeouts()`没有正确设置`timeout`属性的问题
* 修复关闭 tab 时小几率报错问题
* 修复某些情况下元素`size`不准确问题

v3.2.31[​](#v3231 "v3.2.31的直接链接")
---------------------------------

* 页面类增加`user_agent`属性
* `get_src()`方法增加`base64_to_bytes`参数
* 重新设计`find_tabs()`方法
* 使用新版的`DownloadKit`，下载增加追加模式
* `new_tab()`方法的`switch_to`属性默认改为`False`
* `scroll.to_see()`方法的`center`参数默认改为`None`
* `ChromiumOptions`执行`set_argument('--headless')`时能自动使用正确的写法
* `get()`支持 ipv6
* 元素屏幕坐标会乘以像素比再返回
* 问题修复
  + 修复`wait.data_packets()`出现的小概率丢失目标报错
  + 修复当网站 headers 不规范时获取不到编码问题
  + 解决滚动后点击被页面上固定元素遮挡问题
  + 修复某些情况下`back()`后退不准确的情况
  + 修复`'Secure-aa'`和`'Host-'`开头的 cookie 不能设置的问题
  + 修复`WebPage`的`get_cookies()`方法不能获取所有域名的问题
  + 修复`wait.load_start()`不能正确设置超时的问题
  + 修复录屏视频编码一些电脑不支持的问题

v3.2.30[​](#v3230 "v3.2.30的直接链接")
---------------------------------

* 优化抓取数据包逻辑，`wait.data_packets()`删除`targets`参数
* 动作链`type()`可接收`list`和`tuple`
* 浏览器页面对象现在可用 xpath 直接返回文本或注释
* 恢复对 python 3.6 支持
* 完全删除之前声明废弃的方法和属性
* 增加`auto_port`模式可使用端口范围
* 修复`select.by_index()`报错
* 修复`get_session_storage()`报错
* 修复下拉框没有触发`onChange`问题
* 修复`<iframe>`中元素使用`s_ele()`时出现的问题
* 微调`run_js()`逻辑

v3.2.26[​](#v3226 "v3.2.26的直接链接")
---------------------------------

* 新功能
  + 相对定位增加`child()`和`children()`方法
  + 相对定位增加`ele_only`参数
  + 页面对象增加`get_frames()`方法
  + 页面对象增加`wait.new_tab()`方法
  + 页面对象增加`wait.data_packets()`方法
  + `ChromiumPge`增加`find_tabs()`方法
  + 元素对象增加`focus()`方法
  + 元素对象增加`states.is_checked`属性
  + 录屏功能增加非节俭模式和 js 模式
  + 可设置无法点击时抛出异常
  + 元素和动作链增加双击方法
* api 和特性变更
  + `click()`删除`wait_loading`参数
  + `click.at()`增加`count`参数
  + `drag()`和`drag_to()`的`speed`参数改为`duration`
  + `set_headless()`方法适配新版浏览器
  + `ChromiumPage`创建时可只接受端口号
  + `new_tab()`现在会返回新标签页 id
  + `get_frame()`方法增加`timeout`参数，且可接收 id 或 name 为条件
  + `ChromiumFrame`的`wait`属性增加元素特征
  + 录屏功能 api 调整
* 优化和修复
  + 修复同域`ChromiumFrame`没有及时关闭连接问题
  + 改进 cookies 处理逻辑
  + 自动用`'127.0.0.1'`替换`'localhost'`以提高速度
  + 浏览器路径可接受文件夹路径
  + 提高`ChromiumFrame`和查找元素  稳定性
  + 修复`get_local_storage()`和`get_session_storage()`获取所有数据时的问题
  + js 返回字典时能正确解析
  + 修复`get_src()`某情况下`timeout`失效问题
  + 修复`Keys.ENTER`没有正确回车问题

v3.2.19[​](#v3219 "v3.2.19的直接链接")
---------------------------------

* 修改`click()`策略，默认强制模拟点击
* `click()`增加`timeout`参数
* 页面对象添加重试次数和间隔设置
* 不使用 ini 文件时也能适配 Chrome 111 版

v3.2.16[​](#v3216 "v3.2.16的直接链接")
---------------------------------

* 适配 Chrome 111 版
* 修复 cookies 相关问题
* 修复浏览器页面对象`set_headers()`方法无效问题
* 浏览器页面对象`get_cookiees()`方法增加`all_domains`参数
* 优化输入文本前的点击
* `WebPage`的`set_cookies()`方法删除两个参数

v3.2.14[​](#v3214 "v3.2.14的直接链接")
---------------------------------

* 查找元素增加`^`和`$`符号，表示匹配开头和结尾内容
* 页面对象增加`rect.window_state`属性
* 页面对象增加`is_alive`属性
* 元素对象等待增加`enabled()`、`disabled()`、`disable_or_deleted()`方法
* 列表元素增加`by_loc()`、`cancel_by_loc()`、`all()`方法
* `get_cookies()`增加`all_info`参数
* `Session`对象重定向默认为`True`
* 去除对 tldextract 库依赖
* 改善`<iframe>`等待
* 提高启动速度
* 优化元素等待逻辑和下拉菜单逻辑

v3.2.12[​](#v3212 "v3.2.12的直接链接")
---------------------------------

* 支持对异域`<iframe>`元素内的元素截图
* 截图增加`as_base64`参数
* 页面对象增加滚动行为和等待滚动结束设置
* 优化连接浏览器和页面滚动逻辑
* 修复保存图片未正确等待问题
* 修正元素`size`属性返回顺序

v3.2.11[​](#v3211 "v3.2.11的直接链接")
---------------------------------

* 增加录屏功能
* `click()`删除`retry`和`timeout`参数
* `get_src()`和`save()`增加`timeout`参数
* 增加`NoResourceError`异常
* 允许指定使用系统安装的浏览器用户数据文件夹
* 一些其它优化

v3.2.7[​](#v327 "v3.2.7的直接链接")
------------------------------

* `ActionChains`导入路径改为`from DrissionPage.common import ActionChains`
* `Keys`导入路径改为`from DrissionPage.common import Keys`
* 增加`By`类
* 解决启动浏览器时冲突问题
* 修复图片保存和`ChromiumPage`创建时可能遇到的问题

v3.2.5[​](#v325 "v3.2.5的直接链接")
------------------------------

* 浏览器元素增加是否被遮盖属性
* 浏览器元素增加等待被遮盖和等待遮盖取消方法
* 增加`WebPageTab`对象，从`WebPage`生成，可切换模式
* `WebPage`的`get_tab()`方法返回`WebPageTab`对象
* `to_tab()`、`close_tabs()`、`close_other_tabs()`方法可接收标签页对象
* `to_front()`方法移到`set`属性，并增加可指定标签页功能
* 修复用 driver 创建`ChromiumPage`时初始化不正确的问题

v3.2.3[​](#v323 "v3.2.3的直接链接")
------------------------------

* 特性改变

  + `WebPage`取消自动模式切换
  + 找不到元素时返回`NoneElement`，还支持抛出异常
  + 下载默认使用浏览器
  + 元素`wait_ele()`方法取消，改为等待自身状态改变
* 大量整合同类型的 api
* 新增功能

  + 拦截上传控件自动填写路径
  + 优先读取项目路径下的 ini 文件
  + 查找元素增加或语法
  + 新增一批异常
  + 新增命令行工具
  + 页面和元素对象新增一批位置属性
  + `SessionPage`新增一批设置方法
  + 新增几个等待方法
  + 新增`get_frame()`方法
* 优化和修复

  + 对程序底层和业务逻辑进行了重新梳理，优化程序逻辑，大幅增强稳定性
  + 新旧版本完全隔离，新版以后开发可放飞自我，无需担心影响以前用`MixPage`开发的程序
  + 现在会返回开发者能看懂的异常信息
  + 修复页面加载和退出触发弹窗引起的问题
  + 修复`<iframe>`加载时可  能出现的 500 错误
  + 修复异域`<iframe>`点击问题
  + 没有位置和大小信息的元素在获取这些信息时，现在会抛出异常
  + 修复内存没有正确释放的问题
  + 修复点击被固定栏遮挡问题
  + 接管新出现的`<iframe>`会自动等待内容加载
  + 修复`<iframe>`在同域和异域间互相跳转时会卡住的问题
  + 修复`<iframe>`内元素截图出现偏移问题

v3.1.6[​](#v316 "v3.1.6的直接链接")
------------------------------

* `ChromiumPage`添加`latest_tab`属性
* `WebPage`初始化删除`tab_id`参数
* 修复页面未加载完可能获取到空元素的问题
* 修复新标签页重定向时获取文档不正确问题
* 修复使用多标签页或 iframe 时内存未释放问题
* 增强稳定性

v3.1.1[​](#v311 "v3.1.1的直接链接")
------------------------------

* 增强下载功能

  + `ChromiumPage`也可以使用内置下载器下载文件
  + 可拦截并接管浏览器下载任务
  + 新增`download_set`属性对下载参数进行设置
  + 增加`wait_download_begin()`方法
* 改进浏览器启动设置

  + 优化 ini 文件结构
  + 新增`ChromiumOptions`取代`DriverOptions`，完全摆脱对 selenium 的依赖
  + 新增自动分配端口功能
  + 优化`SessionOptions`设计，增加一系列设置参数的方法
  + 改进对用户配置文件的设置
* 对部分代码进 行重构

  + 优化页面对象启动逻辑
  + 优化配置类逻辑
  + 优化项目结构
* 细节

  + 上传文件时支持传入相对路径
* bug 修复

  + 修复`get_tab()`出错问题
  + 修复新浏览器第一次新建标签页时不正确切换的问题
  + 修复关闭当前标签页出错问题
  + 修复改变浏览器窗口大小出错问题

v3.0.34[​](#v3034 "v3.0.34的直接链接")
---------------------------------

* `WebPage`删除`check_page()`方法
* `DriverOptions`和`easy_set`的`set_paths()`增加`browser_path`参数
* `DriverOptions`增加`browser_path`属性
* `ChromiumFrame`现在支持页面滚动
* 改进滚动到元素功能
* 修改`SessionElement`相对定位参数顺序
* `SessionPage`也可以从 ini 文件读取 timeout 设置
* ini 文件中`session_options`增加`timeout`项
* `SessionOptions`增加`timeout`属性、`set_timeout()`方法
* 优化和修复一些问题

v3.0.31[​](#v3031 "v3.0.31的直接链接")
---------------------------------

* `run_script()`、`run_async_script()`更名为`run_js`和`run_async_js()`
* 返回的坐标数据全部转为`int`类型组成的`tuple`
* 修改注释

v3.0.30[​](#v3030 "v3.0.30的直接链接")
---------------------------------

* 元素增加`m_click()`方法
* 动作链增加`type()`、`m_click()`、`r_hold()`、`r_release()`、`m_hold()`、`m_release()`方法
* 动作链的`on_ele`参数可接收文本定位符
* `WebPage`、`SessionPage`、`ChromiumPage`增加`set_headers()`方法

v3.0.28[​](#v3028 "v3.0.28的直接链接")
---------------------------------

* 各种大小、位置信息从`dict`改为用`tuple`返回
* 改进`ChromiumFrame`
* 修复小窗时定位不准问题，修复 iframe 内元素无法获取 s\_ele() 问题
* 增加`wait_loading`方法和参数
* 其它优化和问题修复

v3.0.22[​](#v3022 "v3.0.22的直接链接")
---------------------------------

* `change_mode()`增加`copy_cookies`参数
* 调整`WebPage`生成的元素对象的`prev()`、`next()`、`before()`、`after()`参数顺序
* 修复读取页面时小概率失效问题
* 用存根文件取代类型注解

v3.0.20[​](#v3020 "v3.0.20的直接链接")
---------------------------------

重大更新。推出`WebPage`，重新开发底层逻辑，摆脱对 selenium 的依赖，增强了功能，提升了运行效率。支持 chromium 内核的浏览器（如 chrome 和 edge）。比`MixPage`有以下优点：

* 无 webdriver 特征
* 无需为不同版本的浏览器下载不同的驱动
* 运行速度更快
* 可以跨 iframe 查找元素，无需切  入切出
* 把 iframe 看作普通元素，获取后可直接在其中查找元素，逻辑更清晰
* 可以同时操作浏览器中的多个标签页，即使标签页为非激活状态
* 可以直接读取浏览器缓存来保持图片，无需用 GUI 点击保存
* 可以对整个网页截图，包括视口外的部分（90以上版本浏览器支持）

其它更新：

* 增加`ChromiumTab`和`ChromiumFrame`类用于处理 tab 和 frame 元素
* 新增与`WebPage`配合的动作链接`ActionChains`
* ini 文件和`DriverOption`删除`set_window_rect`属性
* 浏览器启动配置实现对插件的支持
* 浏览器启动配置实现对`experimental_options`的`prefs`属性支持

[上一页

📒 v4.0](/versions/4.0.x)[下一页

📒 v1.5-v2.x](/versions/2x)

* [v3.2.35](#v3235)
* [v3.2.33](#v3233)
* [v3.2.31](#v3231)
* [v3.2.30](#v3230)
* [v3.2.26](#v3226)
* [v3.2.19](#v3219)
* [v3.2.16](#v3216)
* [v3.2.14](#v3214)
* [v3.2.12](#v3212)
* [v3.2.11](#v3211)
* [v3.2.7](#v327)
* [v3.2.5](#v325)
* [v3.2.3](#v323)
* [v3.1.6](#v316)
* [v3.1.1](#v311)
* [v3.0.34](#v3034)
* [v3.0.31](#v3031)
* [v3.0.30](#v3030)
* [v3.0.28](#v3028)
* [v3.0.22](#v3022)
* [v3.0.20](#v3020)

* 📒 v4.0

本页总览

📒 v4.0
======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

v4.0.5.6[​](#v4056 "v4.0.5.6的直接链接")
-----------------------------------

* 优化`auto_port()`逻辑
* `set.cookies()`忽略过期的 cookie
* 修复下载文件可能出现无写入权限报错
* 修复`SessionPage`的 headers 设置问题
* 修复链接以`'./'`开头时`ele.link`获取不准确的问题
* 修复异域`<iframe>`跳转到同域时的问题

---

v4.0.5.2[​](#v4052 "v4.0.5.2的直接链接")
-----------------------------------

* 增加视觉相对定位语法
* 改进元素结果列表筛选功能
* `wait.has_rect()`和`wait.covered()`返回具体信息
* 删除元素获取任意属性语法
* 删除之前声明废弃的参数、方法和属性
* 修复`states.is_alive()`和`wait.deleted()`问题

---

v4.0.4.25[​](#v40425 "v4.0.4.25的直接链接")
--------------------------------------

* 支持对`eles()`返回的列表进行筛选
* `DrissionPage.common`新增`get_eles()`方法，可接收多个定位符获取多个元素
* `input()`输入前会自动等待元素可点击
* `set.cookies()`接收`str`形式 cookies 时，只支持用`';'`做分隔符
* 修复监听一个报错

---

v4.0.4.23[​](#v40423 "v4.0.4.23的直接链接")
--------------------------------------

* 元素增加`states.is_clickable`属性和`wait.clickable()`、`set.style()`方法
* `tree()`增加`text`、`show_js`、`show_css`参数
* `wait.stop_moving()`参数顺序改变
* `tab_ids`属性不再屏蔽插件标签页
* 修复`SessionPage`的`get()`方法访问本地中文路径问题
* 等待元素时可抛出页面断开异常

---

v4.0.4.22[​](#v40422 "v4.0.4.22的直接链接")
--------------------------------------

* 动作链`scroll()`方法参数位置变化
* 页面对象的`save()`方法可根据后缀自动判断类型
* 中键单击返回 Tab 对象
* `tab_ids`属性忽略插件 tab
* 优化 cookies 设置逻辑
* Frame 对象初始化时不再等待 url 变化
* 修复全局代理时无法连接浏览器的问题
* 修复截图文件名过长时的问题
* 修复带 html 节点的 shadow root 获取不到子元素问题
* 降低失去元素报错可能性

---

v4.0.4.21[​](#v40421 "v4.0.4.21的直接链接")
--------------------------------------

* `add_ele()`的`outerHTML`参数改为`html_or_info`，可新增不插入到 DOM 的元素
* `wait.ele_loaded()`改成`wait.eles_loaded()`，可等待多个元素全部或任一个加载
* 取消无界面 Linux 自动无头功能
* 调整`quit()`逻辑
* 修复 prompt 无法输入的问题
* 修复`WebPageTab`的`close()`报错问题
* 修复下拉列表已选中元素再点击会取消的问题
* 修复`run_js()`无法添加`dict`参数问题
* 修复`set.cookies()`的一个问题

---

v4.0.4.17[​](#v40417 "v4.0.4.17的直接链接")
--------------------------------------

* Page 对象的`set.auto_handle_alert()`增加`all_tabs`参数
* 修复`ele.text`速度慢的问题
* 修复在未访问网页时设置`'__Host-'`开头的 cookie 时出现的问题

---

v4.0.4.16[​](#v40416 "v4.0.4.16的直接链接")
--------------------------------------

* `ChromiumPage`增加`browser_version`属性
* `DataPacket.request`增加`cookies`属性
* `wait_slient()`方法增加`limit`参数
* `Keys`增加`CTRL_A`等 6 个组合键
* 元素的`save()`方法增加`rename`参数
* 元素的`input()`方法的`clear`参数默认值改为`False`
* 动作链`type()`可接收键盘以外的字符
* `get_tab()`和`get_tabs()`默认获取普通 tab
* 修复动作链`wait()`问题

---

v4.0.4.13[​](#v40413 "v4.0.4.13的直接链接")
--------------------------------------

* 动作链`wait()`增加随机功能
* 当 tab 设置不为单例时，`latest_tab`返回 tab id
* 修复新标签页重复创建连接问题
* 修复等待新 tab 不正确问题

---

v4.0.4.12[​](#v40412 "v4.0.4.12的直接链接")
--------------------------------------

* 元素增加获取任意属性方式
* 调整`quit()`逻辑
* 优化`set.cookies()`逻辑
* 修复`get_tab()`问题
* 修复`set_flags()`特定情况下出现的问题
* 修复某`WebPage`在些情况下`get_tab()`时出现的问题

---

v4.0.4.9[​](#v4049 "v4.0.4.9的直接链接")
-----------------------------------

* `get_tab()`方法增加几个参数
* `find_tabs()`方法改为`get_tabs()`，且改为默认返回标签页对象
* 设置`headers`时可接收从浏览器复制的文本格式
* `common`路径增加`from_selenium()`和`from_playwright()`方法
* `latest_tab`改为返回标签页对象
* `tabs`属性改为`tab_ids`
* 修复从文本生成的静态元素`link`属性报错问题
* 修复保存的 mhtml 文件无法被浏览器解析问题

---

v4.0.4.8[​](#v4048 "v4.0.4.8的直接链接")
-----------------------------------

* 增加`click.for_new_tab()`方法
* `wait()`方法增加`scope`参数，可等待指定随机时间
* `set.upload_files()`和`click.to_upload()`方法支持接收`Path`类型
* `click.to_download()`增加`timeout`参数
* 元素对象`page`属性改为指向总体的 Page对 象，增加`owner`属性
* 完善找 chrome 路径逻辑
* 调整`quit()`逻辑
* 修复处理有些下拉列表时报错问题
* 修复页面用 css 查找元素时会找到文本的问题
* 修复用 css 在元素下获取多个子元素时数量不正确问题
* 修复在 ini 文件设置下载路径时报错问题
* 修复`run_async_js()`报错问题
* 修复`reconnect()`报错问题

---

v4.0.4.5[​](#v4045 "v4.0.4.5的直接链接")
-----------------------------------

* 截图左上和右下参数可只接收其中一个
* 配置对象`save()`可生成不存在的路径
* 增加`click.to_upload()`和`click.to_download()`方法
* `DrissionPage.common`增加`tree()`方法
* 去除`upload()`方法
* 修复`handle_alert()`
* 修复一个 js 结果解析问题
* 修复命令行问题

---

v4.0.4.3[​](#v4043 "v4.0.4.3的直接链接")
-----------------------------------

* `ChromiumOptions`增加`clear_arguments()`、`clear_prefs()`和`clear_flags()`方法
* 浏览器页面对象增加`upload()`方法
* 浏览器页面对象增加`add_ele()`方法
* `run_js()`方法可读取文件
* `click.multiple()`改为`click.multi()`
* 修复旧版 python 中`get()`报错问题

---

v4.0.4.1[​](#v4041 "v4.0.4.1的直接链接")
-----------------------------------

* 打包程序即使不带上 ini 文件也不会报错
* `SessionOptions`增加`clear_headers()`方法
* `Settings`增加`cdp_timeout`属性
* `prop()`改成`property()`，参数改为`name`
* `get_cookies()`改为`cookies()`
* `get_src()`改为`src()`
* 删除`cookies`属性
* `get_session_storage()`、`get_local_storage()`改为`session_storage()`、`local_storage()`
* `pageLoad`改成`page_load`
* `set_a_header()`、`remove_a_header()`、`set.header()`、`set.attr()`的参数改为`name`
* 元素增加`value`属性和`set.value()`方法
* `loc_or_ele`、`loc_or_str`等参数改为`locator`
* 提高截图 jpg 格式画质
* 修复 s 模式`timeout`参数失效问题
* 修复`wait.has_rect()`等出现的问题
* 修复找不到浏览器路径时报 ini 错误问题
* 增加一些提示

---

v4.0.3.4[​](#v4034 "v4.0.3.4的直接链接")
-----------------------------------

此版本对项目进行了大量重构，新增了不少功能，改善了运行逻辑，优化了项目结构，解决了很多以前积累的问题。对比旧版本有质的提高。

但同时很多 api 产生了变化，不能完全兼容旧版本。

* 改进抓包功能
  + 页面对象新增`listen`属性，弃用`FlowViewer`
  + 删除`wait.set_targets()`删除
  + 删除`wait.stop_listening()`方法
  + 删除`wait.data_packets()`方法
  + `DrissionPage.common`路径删除`FlowViewer`
  + 用`listen.set_start()`和`listen.stop()`启动和停止监听
  + 用`listen.wait()`阻塞等待数据包
  + 用`listen.steps()`同步获取监听结果
  + 增加`listen.wait_silent()`等待所有请求完成（包含 targets 以外的）
  + 监听结果结构优化，request 和 response 数据分开存放
* 重构连接逻辑
  + 页面对象`page_load_strategy`属性改名为`load_mode`
  + `set.load_strategy`改为`set.load_mode`
  + `get()`方法的`timeout`参数现在可覆盖整个过程
  + `timeout`参数对非`get()`方法触发的加载（如点击链接）也能生效
  + `SessionPage`和`WebPage`的 s 模式，如收到空数据，也会重试
  + `SessionPage`的`get()`方法可以指向本地文件
  + 新的`none`加载模式
* 改进下载管理功能
  + 页面对象删除`download_set`属性
  + 增加`set.download_path()`方法
  + 增加`set.download_file_name()`方法
  + Tab 对象和 Frame 对象也支持`download()`方法
  + 每个 Tab 对象可单独设置下载路径和重命名文件名
  + 可拦截浏览器下载任务并获取其信息
  + 可取消浏览器下载任务、获取下载进度、等待任务完成
  + 可设置遇到文件夹已存在时的处理方式
  + 默认不启用浏览器下载任务管理
* 页面对象改进
  + `ChromiumPage`和`WebPage`改为固定单例
  + `get_tab()`获取的 Tab 对象默认单例，可用`Settings`设置允许多例
  + 浏览器页面对象启动时不再接收`ChromiumDriver`对象
  + `WebPage`对象的`driver_or_options`参数 改名为`chromium_options`
  + `ChromiumPage`对象的`addr_driver_opts`参数改名为`addr_or_opts`
  + 页面对象内置动作链
  + `ready_state`、`is_loading`、`is_alive`属性合并到`states`属性中
  + 页面对象增加`raw_data`参数，s 模式下返回原始数据
  + 所有页面对象增加`close()`方法，`SessionPage`用于关闭连接，浏览器页面对象用于关闭标签页
  + 浏览器页面对象增加`wait()`方法，用于等待若干秒
  + 浏览器页面对象增加`wait.ele_loaded()`方法，等待元素加载到DOM
  + 浏览器页面对象增加`wait.title_change()`和`wait.url_change()`方法，用于等待 title 和 url 变化
  + 浏览器页面对象增加`wait.alert_closed()`方法，用于等待弹窗被手动关闭
  + 浏览器页面对象增加`set.blocked_urls()`方法，可设置忽略的连接
  + Tab 和 Page 对象增加`disconnect()`、`reconnect()`和`save()`方法
  + Tab 和 Page 对象增加`add_init_js()`和`remove_init_js()`方法
  + `wait.ele_delete()`方法改为`wait.ele_deleted()`
  + `wait.ele_display()`方法改为`wait.ele_displayed()`
  + `wait.load_complete()`方法改为`wait.doc_loaded()`
  + `quit()`方法增加`force`参数，可强制关闭浏览器进程
  + `ChromiumFrame`增加`ract`属性
  + `ChromiumFrame`的`frame_size`属性改为`rect.size`
  + 优化`SessionPage`和`WebPage`s 模式访问速度
  + `WebPage`在 d 模式时，`post()`返回`Response`对象
* 标签页管理改进
  + 删除`to_tab()`方法
  + 删除`to_main_tab()`、`set.main_tab()`方法
  + 删除`main_tab`属性
  + `new_tab()`方法删除`switch_to`参数
  + `new_tab()`方法增加`new_window`、`background`、`new_context`参数
  + `rect.borwser_size`改为`rect.window_size`
  + `rect.borwser_location`改为`rect.window_location`
  + `set.window.maximized()`改为`set.window.max()`
  + `set.window.minimized()`改为`set.window.mini()`
  + `set.window.fullscreen()`改为`set.window.full()`
  + Tab 对象增加`set.activate()`、`close()`、`handle_alert()`、`states.has_alert`
  + `get_tab()`的`tab_id`参数改为`id_or_num`，可接收序号
* cookies 设置改进
  + `set.cookies()`可接收单个 cookie
  + 增加`set.cookies.clear()`方法用于清除 cookies
  + 增加`set.cookies.remove()`方法用于删除一个 cookie 项
* 元素相关改进
  + 查找元素增加`@!`语法
  + 删除`@@-`和`@|-`语法
  + `ele()`和`s_ele()`增加`index`参数，可指定获取第几个
  + 相对定位第一个参数支持接收序号
  + 位置和大小
    - 删除`size`、`location`、`locations`属性，新增`rect`属性
    - 旧版中`loactions.****`的属性改为`rect.****`
    - 大小和位置信息，从`int`类型改为`float`类型
    - 增加`states.has_rect`属性，返回元素是否拥有大小和位置
    - 增加`states.is_whole_in_viewport`属性，返回元素是否整个都在视口内
  + 点击改进
    - `click()`增加`wait_stop`参数，默认等待元素运动停止再点击
    - `click()`默认等待元素运动停止再执行点击
    - `click.twice()`改为`click.multiple()`
  + 查找元素失败显示细节
  + 可设置查找失败元素返回默认值
  + 增加`wait.stop_moving()`方法，可等待移动结束
  + 增加`wait()`方法，等待若干秒
  + 增加`check()`方法，可选中或取消选中元素
  + 滚动添加`to_center()`方式，可滚动到视口中央
  + 增加`select.by_option()`和`select.cancel_by_option()`方法，可选取列表项元素
  + 增加`states.has_rect`属性
  + 添加`states.is_whole_in_viewport`属性，判断是否整个都在视口中
  + 元素被覆盖时，`states.is_covered`属性返回覆盖元素的 id
  + `input()`方法增加`by_js`参数
  + `save()`的`rename`参数改为`name`
  + `get_src()`支持 blob 类型
  + `css_path`获取的路径更精确
  + 相对定位的`timeout`参数默认改为`None`
  + `wait.delete()`方法改为`wait.deleted()`
  + `wait.disabled_or_delete()`方法改为`wait.disabled_or_deleted()`
  + `wait.display()`方法改为`wait.displayed()`
  + 可用`==`比较两个元素
  + 查找元素速度提高
* 启动配置改进
  + 删除 easy\_set 方法
  + 启动或接管浏览器时，可自动关闭弹出的隐私声明
  + 在无界面系统启动浏览器时，自动使用无头，可用`set_headless(False)`禁用
  + 当`set_headless(False)`但接管了无头浏览器，将关闭并启动新的有头浏览器
  + `auto_port()`方法支持多线程
  + ini 文件
    - `chrome_options`类改为`chromium_options`
    - `binary_location`项改为`browser_path`
    - `page_load_strategy`项改为`load_mode`
    - `debugger_address`项改为`address`
    - `arguments`项删除`'--remote-allow-origins=*'`参数
    - `arguments`项增加`'--no-default-browser-check'`、`'--disable-suggestions-ui'`、`'--disable-popup-blocking'`、`'--hide-crash-restore-bubble'`、`'--disable-features=PrivacySandboxSettings4'`参数
    - `paths`类增加`tmp_paht`项
    - 删除`experimental_options`项
    - `chrome_options`类增加`prefs`、`flags`、`existing_only`项
    - 增加`others`类，包含`retry_times`和`retry_interval`项
  + `ChromiumOptions`
    - 增加`set_flag()`和`clear_flags_in_file()`，用于设置实验项
    - 增加`existing_only()`方法和`is_existing_only`属性，可指定只接管浏览器而不自动启动新的
    - 增加`ignore_certificate_errors()`方法，可忽略证书报错
    - 增加`retry_times`、`retry_interval`属性和`set_retry()`方法，可设置重试参数
    - 增加`incognito()`方法，可设置无痕模式
    - 增加`set_tmp_path()`方法
    - 增加`tmp_path`和`is_auto_port`属性
    - `auto_port()`增加`tmp_path`参数
    - `set_paths()`
      方法拆分成`set_browser_path()`、`set_local_port()`、`set_address()`、`set_download_path()`、`set_user_data_path()`、`set_cache_path()`
      方法
    - `set_page_load_strategy()`改成`set_load_mode()`
    - `set_headless()`改成`headless()`
    - `set_no_imgs()`改成`no_imgs()`
    - `set_no_js()`改成`no_js()`
    - `set_mute()`改成`mute()`
    - `debugger_address`改成`address`
  + `SessionOptions`
    - `SessionOptions`的`set_paths()`方法改为`set_download_path()`
    - 增加`retry_times`、`retry_interval`属性和`set_retry()`方法，可设置重试参数
* 其它
  + 删除 2.x 代码
  + `handle_alert()`方法增加`next_one`参数，可处理下一个出现的弹窗
  + 浏览器页面对象增加`set.auto_handle_alert()`方法，可设置自动处理弹窗
  + `SessionPage`增加`set.encoding()`方法和`encoding`属性
  + `<option>` 元素可以接受点击，操作更符合直觉
  + `run_js()`、`run_js_loaded()`、`run_async_js()`方法增加`timeout`参数
  + `run_async_js()`删除`timeout`参数
  + `timeouts`的`implicit`改成`base`
  + `ActionChains`改成`Actions`
  + 动作链的移动方法增加`duration`参数
  + 动作链增加`input()`方法
  + 动作链`key_down()`和`key_up()`方法可接收按键名称文本
  + 动作链`type()`方法`text`参数改为`keys`
  + `get_screenshot()`方法增加`name`属性，可指定文件名
  + 元素的`get_screenshot()`方法增加`scroll_to_center`参数，截图前先滚动到页面正中
  + `wait.new_tab()`方法成功时返回新标签页 id
  + `tabs`不包含 F12 的窗口
  + `DrissionPage.common`路径增加`wait_until()`方法，支持自定义组合等待条件
  + `DrissionPage.common`路径增加`get_blob()`方法
  + 异常变化
    - `CallMethodError`改为`CDPError`
    - `ElementLossError`改为`ElementLostError`
    - `ContextLossError`改为`ContextLostError`
    - `TabClosedError`改为`PageDisconnectedError`
    - 增加`WaitTimeoutError`
    - 增加`GetDocumentError`
    - 增加`WrongURLError`
    - 增加`StorageError`
    - 增加`CookieFormatError`
    - 增加`TargetNotFoundError`
  + Settings 变化
    - 增加`singleton_tab_obj`，设置 Tab 对象是否允许多例
    - `raise_ele_not_found`改为`raise_when_ele_not_found`
    - `raise_click_failed`改为`raise_when_click_failed`
* 优化
  + MAC 和 Linux 系统添加默认浏览器路径
  + 全面重构对象启动和运行逻辑，大幅提高稳定性
  + 接管或启动浏览器不再要求`--remote-allow-orignins`参数
  + 所有涉及循环的代码都加上超时设施，杜绝卡死
  + 对`ChromiumFrame`进行完全重构，提高稳定性
  + 调整项目结构
* 问题修复
  + 修复网络连接极不稳定时获取文档失败问题
  + 修复相对定位`timeout`失效问题
  + 修复 shadow root 内定位元素可能偏差问题
  + 修复异域`ChromiumFrame`内部元素无法获取屏幕坐标的问题
  + 修复相对路径插件加载失败的问题
  + 所有循环增加超时设置，避免出现卡死
  + 修复元素截图时窗口外部分空白问题
  + 修复 Tab 没有继承 Page 下载路径的问题
  + 修复 `<iframe>` 内元素获取 href 属性错误问题
  + 修复 cookie 设置 expires 时的问题

[上一页

📒 v4.1](/versions/4.1.x)[下一页

📒 v3.x](/versions/3x)

* [v4.0.5.6](#v4056)
* [v4.0.5.2](#v4052)
* [v4.0.4.25](#v40425)
* [v4.0.4.23](#v40423)
* [v4.0.4.22](#v40422)
* [v4.0.4.21](#v40421)
* [v4.0.4.17](#v40417)
* [v4.0.4.16](#v40416)
* [v4.0.4.13](#v40413)
* [v4.0.4.12](#v40412)
* [v4.0.4.9](#v4049)
* [v4.0.4.8](#v4048)
* [v4.0.4.5](#v4045)
* [v4.0.4.3](#v4043)
* [v4.0.4.1](#v4041)
* [v4.0.3.4](#v4034)

* 📒 v4.1

本页总览

📒 v4.1
======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

v4.1.1.2[​](#v4112 "v4.1.1.2的直接链接")
-----------------------------------

* 支持 ws 方式连接浏览器
* 动作链增加`drag_in()`方法
* 插件改用临时方式加载，以适配 135 版之后的浏览器
* 隐藏`ChromiumOptions`的`set_paths()`方法，未来将弃用
* 其它小问题修复

---

v4.1.0.18[​](#v41018 "v4.1.0.18的直接链接")
--------------------------------------

* 修复`click.middle()`问题
* 修复获取元素`link`属性不正确问题

---

v4.1.0.17[​](#v41017 "v4.1.0.17的直接链接")
--------------------------------------

* 元素增加`child_count`属性
* `Settings`每个属性都增加`set_****()`方法
* 增加英文版报错和提示
* 增加`LocatorError`和`UnknownError`异常
* `ShadowRoot`能够返回 xpath 的文本或数字结果
* `SessionElement`的`attrs`返回未处理链接属性
* `WrongURLError`改为`IncorrectURLError`
* `suffixes_list_path`改为`suffixes_list`
* `ChromiumElement`的`attr()`方法`attr`参数改为`name`
* 调整部分报错类型
* 修复无痕模式`new_tab()`的一个问题
* 修复某些时候执行 js 报错问题
* 修复设置 suffixes\_list 问题
* 修复一个`SessionPage`访问 Linux 本地路径时的问题
* 修复在`iframe`内的元素屏幕坐标不正确的问题

---

v4.1.0.14[​](#v41014 "v4.1.0.14的直接链接")
--------------------------------------

* `screencast.stop()`增加`suffix`和`coding`参数
* 修复一个抛出异常设置问题
* 修复一个 js 返回结果解析问题

---

v4.1.0.13[​](#v41013 "v4.1.0.13的直接链接")
--------------------------------------

* `ChromiumFrame`添加`style()`方法
* 指定 DownloadKit 2.0.7 或以上版本
* 修改 LICENSE
* 提高运行速度
* 修复 js 录像报错问题

---

v4.1.0.12[​](#v41012 "v4.1.0.12的直接链接")
--------------------------------------

* 动作链`type()`方法增加`interval`参数
* Page 对象加上几种浏览器状态
* 增加`Settings.suffixes_list_path`，用于设置离线域名后缀列表文件
* 支持离线运行
* DownloadKit 指定 2.0.5 版本，`download()`的`goal_path`改为`save_path`
* `Mission`对象`path`属性改为`folder`，增加`tmp_path`属性
* 优化`css_path`
* 优化等待新标签页逻辑
* 优化关闭标签页逻辑
* 接管来自 selenium 和 playwright 的浏览器时忽略无头设置
* 增加`Settings.browser_connect_timeout`属性
* `remove_attr()`返回元素自身
* `select`各种方法返回元素本身，找不到项时报错
* 指定 tldextract 版本需大于等于 3.4.4
* 优化关闭标签页逻辑
* 点击产生的新标签页下载任务可用原标签页等待
* 修复下载路径设置问题
* 修复`new_tab()`时浏览器关闭导致的卡住
* 修复一个 headers 设置问题
* 修复多线程关闭标签页时可能报错问题
* 修复无法处理连续出现的弹出框的问题
* 修复新建标签页可能出现的问题

---

v4.1.0.7[​](#v4107 "v4.1.0.7的直接链接")
-----------------------------------

* `DataPacket`对象增加`request.params`属性
* `DataPacket`对象 headers 补充完整
* `MixTab`和`WebPage`的`s_ele()`补上`timeout`参数
* `wait.has_rect()`、`wait.covered()`成功时返回调用者
* 元素列表切片时也返回列表对象
* 修复`new_tab()`访客模式下不输入`url`参数时报错问题
* 修复`get_tab()`找不到指定标签页对象时报错问题
* 修复某些网站`back()`后卡住问题
* 修复`xpath`属性指向元素不唯一问题

---

v4.1.0.5[​](#v4105 "v4.1.0.5的直接链接")
-----------------------------------

* 引入`Chromium`对象用于代表浏览器
* `WebPageTab`改名为`MixTab`
* `SessionPage`、`ChromiumPage`和`WebPage`初始化时删除`timeout`提示，以后会废弃
* `activate_tab()`取代`set.tab_to_front()`
* Frame 对象增加`set.property()`、`set.style()`、`link`
* 元素对象增加`get_frame()`方法
* 所有对象增加`find()`方法，用于同时匹配多个定位符
* `quit()`增加`del_data`参数
* Tab 对象的`close()`方法增加`others`参数
* `cookies()`删除`as_dict`参数，增加`as_dict()`、`as_json`和`as_str()`方法
* 浏览器页面和元素对象的`s_ele()`和`s_eles()`方法增加`tiemout`参数
* 浏览器页面和元素对象增加`rect.scroll_position`属性
* 动作链删除`db_click()`，各点击方法增加`times`参数
* `wait.new_tab()`增加`curr_tab`参数
* 滚动增加`scroll()`方法
* 部分等待方法会返回调用者，方便链式操作
* `ChromiumOptions`增加`new_env()`方法，ini 文件增加`new_env`参数，用于指定必须用新环境
* `ChromiumOptions`增加`is_headless`属性
* `parent()`和 shadow-root 内查找方法增加`timeout`参数
* 元素对象各种动作返回元素本身，便于链式操作
* 元素对象增加`timeout`属性
* 页面对象增加`console`属性，可读取控制台信息
* 打印`NoneElement`改成详细信息
* `wait.alert_closed()`增加`timeout`参数
* `auto_port()`方法删除`tmp_path`参数
* `src()`方法可获取`<link>`指向的文件内容
* 录像改为 H.265 编码
* `shadow_root`属性增加等待附加到元素（超时 10 秒）
* `set.cookies()`  忽略过期 cookie
* `ChromiumFrame`对象默认改为单例
* `timeout`属性不再接受赋值
* 优化连接浏览器失败报错
* 优化`css_path`
* 修复`new_tab()`在访客模式和隐私模式的问题
* 修复 Frame 对象滚动填入`tuple`定位符报错问题
* 修复`states.is_displayed`有些情况下不正确问题
* 修复元素`link`属性不正确的问题
* 修复 shadow-root 内用 css 找元素的一个问题
* 修复异域`<iframe>`内元素屏幕坐标不准问题
* 修复`new_tab=True`时下载路径不正确问题
* 修复`attr()`填入大写字母无法获取问题

[下一页

📒 v4.0](/versions/4.0.x)

* [v4.1.1.2](#v4112)
* [v4.1.0.18](#v41018)
* [v4.1.0.17](#v41017)
* [v4.1.0.14](#v41014)
* [v4.1.0.13](#v41013)
* [v4.1.0.12](#v41012)
* [v4.1.0.7](#v4107)
* [v4.1.0.5](#v4105)

* 📆 下一步计划

本页总览

📆 下一步计划
=======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

✅️ 下一步计划[​](#️-下一步计划 "✅️ 下一步计划的直接链接")
-------------------------------------

欢迎投喂加速😘

* 支持跨标签页监听数据包
* 支持独立 js 环境
* 模拟手机环境，支持触摸操作
* 支持随时修改浏览器代理
* 研究远程控制浏览器方案
* 支持 websocket 抓包
* 支持操作插件小窗
* 移动和点击时显示鼠标位置

---

✅️ 加快开发速度[​](#️-加快开发速度 "✅️ 加快开发速度的直接链接")
----------------------------------------

开源项目的推进有赖你的支持。欢迎投喂加速。

![](/assets/images/code-764dd10f15f131f48c0319d0eaf8f4fd.jpg)

[上一页

📒 v0.x-v1.4](/versions/1x)[下一页

📖 关于项目](/versions/statement)

* [✅️ 下一步计划](#️-下一步计划)
* [✅️ 加快开发速度](#️-加快开发速度)

* 📖 关于项目

本页总览

📖 关于项目
======

[![](/img/ad.png)](https://b23.tv/w62Tiqd)

✅️ 缘起[​](#️-缘起 "✅️ 缘起的直接链接")
----------------------------

从前，曾经有那么一个古怪无比的系统。

它业务逻辑相当奇葩，界面设计无法理解，运行速度时快时慢，错误处理若有若无。

用过的人都怨声载道。

很不幸，这个系统每年要开展大规模的业务。

更不幸的，本人就是这个系统的管理员之一。

每当业务开展，案头的电话每天都被打爆。

为了处理收集回来的数据，还要人工整理一个星期。

简单来说，这是个让生产效率倒退的工具。

忍无可忍，本人怒学自动化，以将自己从苦逼的繁重劳动中拯救出来。

---

✅️ 入门和发展[​](#️-入门和发展 "✅️ 入门和发展的直接链接")
-------------------------------------

网上翻了一下，selenium 是当时最热门的网页自动化工具，于是边学边用，就此开始接触自动化之路。

一开始，感觉 selenium 简直神奇，十几行语句就能够搞定以前要干大半天的活。

随着了解的深入，逐渐觉得 selenium 就像毛坯房，它只提供了最基础的工具，但要用起来，还要自己进行相当规模的封装。于是就学会了 POM 模式，开始捣鼓自己的 page 对象。

但是，新手入门，各种奇奇怪怪的报错让人无所适从，各种不稳定情况不知如何入处置。甚至还有一些自带的坑难以处置。这段时间，踩了无数的坑，也花了无数精力，填了无数的坑。

慢慢地，自己封装的工具逐渐成熟起来，随着使用场景的增加，也有了更多的需求。

---

✅️ 诞生[​](#️-诞生 "✅️ 诞生的直接链接")
----------------------------

根据经验和需求，我总结出以下需求：

首先，selenium 的语句实在是太啰嗦了，隐式等待适用范围太窄，显式等待语句复杂得过分，查找元素语句冗长，链式操作显得很难看等等，令我这个极简主义者难以忍受。

其次，查找元素的方式太不友好了，这是最常用的操作，但经常写得又长又臭，我希望创造一套简洁高效的查找元素语法。

还有，我希望把浏览器和 requests 组合起来，各取所长，兼顾写得快和跑得快。

最后，我已经封装了一批好用的方法，填了 selenium 本身不少坑，我希望能够走到哪里都能方便地用到自己顺手的工具。

所以，就有了这个库的诞生。

Drission 这是本人自创的词。是 Driver 前半部分和 Session 后半部分的组合体。

因为 selenium 控制浏览器的对象叫`WebDriver`，requests 用于收发数据包的对象叫`Session`，Drission 就是对它们组合的一个尝试。

而 Page，表示本库以页面为单位。

---

✅️ 版本迭代[​](#️-版本迭代 "✅️ 版本迭代的直接链接")
----------------------------------

本人虽有编程和前端的基础，但 python 也是自学，属于摸着石头过河。

好在项目驱动的学习方式进展确实快，项目需要什么，就去学什么知识，就像拼图一样，慢慢地把各种知识补全了。

开始时还是太多东西不懂，就去找现成的库来用。v0.x 至 v1.4 版本是基于 selenium 和 requests-html 制作，前者负责控制浏览器部分，后者负责收发数据包部分。

这一阶段，实现了控制浏览器和收发数据包两者 api 的统一，cookies 的互通，建立了基本的使用逻辑。

但用 request-html 作为底层毕竟有点太重了，在逐渐摸清了它的运行原理之后，本人对这部分底层代码用 requests 和 lxml 进行了重构。来到了第二阶段。

第二阶段是 v1.5 至 v2.x 版本，控制浏览器部分依然基于 selenium，收发数据包和解析功能则完全自己开发。这时跑起来感觉轻松了许多，也增加了更多实用功能和优化。但瓶颈转移到 selenium 这边。

用得越多，了解得越深，对 selenium 的不满就越发增加。selenium 受限于 chromedriver，我的很多想法都无法实现。比如无法对整个网页截图；比如一个 Driver 对象同一时间只能操作一个标签页或页面框架，来回切换标签页的时候，原来已获取到的元素会失效；比如要为不同版本浏览器下载对应 chromedriver，浏览器自动升级最新版时可能没有新驱动而无法使用等等。

还有最重要的一点，近年来我们国家被老美各种打压，本人心中早就憋了一股气，也很想为国内开源事业贡献一点点微薄的力量。

于是来到了第三阶段。

经过 2-3 年的使用，本人踩了足够多的坑，对自动化已颇有些心得，抱着试一试的心态，大胆地迈出了自研底层的一步。在 3.x 版本，DrissionPage 完全放弃了对 selenium 的依赖，自己对底层进行了重构。

真是自研一念起，刹那天地宽。摆脱了 chromedriver 框架的制约，顿时感觉自由的气息扑面而来。从此，DrissionPage 不仅比 selenium 跑得快，还可以实现各种黑科技。这些，用过的各位应该有所体会，这里就不再啰嗦。

经过 3.x 版本的尝试，作者逐渐掌握了框架的窍门，在 4.0、4.1 版本，对结构进行了大幅重构，使作品趋于成熟。

对了，值得一提的是，DrissionPage 有一个副作用。它竟然能够通过 cloudflare、Google 等人机检测工具，这个是作者都没有想到的。也许是自己写的小众玩意，这些大厂还不认识吧。

---

✅️ 碎碎念[​](#️-碎碎念 "✅️ 碎碎念的直接链接")
-------------------------------

不知为何，3.x 之后，几年来默默无闻的 DrissionPage 竟然有点热闹起来了。Gitee 给发了个 GVP。GitHub 的星星直线上升。有点惊喜，感谢大家的厚爱。

其实作者是个很佛系的开发者，本职并非开发，写库只是业余爱好。事实上，随着功能日渐丰富，许多功能自己都用不到，持续开发更多是来自于兴趣。自己细心雕刻的一个作品，希望能够尽善尽美。自己写的代码在世界上运行着，就像延续了我的生命。更重要的，是感觉自己给国内软件业做了一点点贡献，感到有意义。

但是，自动化的软件往往是一把双刃剑。在这里作者得叠一下甲。请勿将 DrissionPage 应用到任何可能会违反法律规定和道德约束的工作中。请友善使用 DrissionPage，遵守蜘蛛协议，不要将 DrissionPage 用于任何非法用途。如您选择使用 DrissionPage 即代表您遵守此协议，作者不承担任何由于您违反此协议带来任何的法律风险和损失，一切后果由您承担。

结尾回收一下开头，那个曾经令我深恶痛绝的系统，开发团队非常负责，虽然起步不好，但他们积极参与到业务中，不断也迭代产品，经过几大版本的更新，现在那个系  统已经好用得不得了。因此，本人对中国软件充满希望，未来会更好。

[上一页

📆 下一步计划](/versions/next)

* [✅️ 缘起](#️-缘起)
* [✅️ 入门和发展](#️-入门和发展)
* [✅️ 诞生](#️-诞生)
* [✅️ 版本迭代](#️-版本迭代)
* [✅️ 碎碎念](#️-碎碎念)

