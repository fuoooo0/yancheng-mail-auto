# 邮箱账单自动入库

把采货侠 / 拍机堂发到网易邮箱的账单附件，自动写入「雁城库存管理」的待入库。

- 只处理指定发件人 / 主题关键词 + `.xlsx` / `.csv` 附件
- 广告、通知、促销邮件不会进待入库
- 拍机堂按「物品编号」去重（同一订单多台都能进），其他账单优先按 IMEI 去重
- 自动跳过重复单号 / 重复 IMEI、退款单、取消单
- 使用 GitHub Actions 定时跑，不需要服务器，不需要电脑常开

## 一、先准备 5 个值

| Secret 名称 | 填什么 |
|---|---|
| `NETEASE_USER` | 你的完整网易邮箱，例如 `xxx@163.com` |
| `NETEASE_PASS` | 网易邮箱「客户端授权码」，不是登录密码 |
| `SUPABASE_URL` | Supabase 项目地址，例如 `https://xxxx.supabase.co` |
| `SUPABASE_KEY` | 和 ERP 云同步里填的同一个 anon key |
| `SHOP_CODE` | ERP 设置里填的店铺 code |

可选：

| Secret 名称 | 填什么 |
|---|---|
| `NETEASE_HOST` | 163 邮箱不用填；126 邮箱填 `imap.126.com` |
| `BILL_KEYWORDS` | 默认 `采货侠,拍机堂,转转`，一般不用改 |

## 二、建 GitHub 仓库

1. 浏览器打开 `https://github.com`，登录。
2. 右上角 `+` → `New repository`。
3. 仓库名填 `yancheng-mail-auto`，选 **Public**，**不要勾 Add a README**，点 `Create repository`。
4. 进入仓库后，点上方 `Add file` → `Upload files`，上传：
   - `mail_to_inbox.py`
   - `requirements.txt`
   - `README.md`
5. 再点 `Add file` → `Create new file`，文件名填：
   `.github/workflows/mail.yml`
   把本包里的 `.github/workflows/mail.yml` 内容整段粘进去，点 `Commit changes`。

> GitHub 网页上传不支持直接保留文件夹层级，所以 workflow 文件用「新建文件」的方式放到 `.github/workflows/` 下最稳。

## 三、填 Secrets

1. 在仓库页面右上角点 `Settings`（不是个人头像里的 Settings）。
2. 左侧栏拉到最下面，点 `Secrets and variables` → `Actions`。
3. 点绿色按钮 `New repository secret`。
4. 按上面的 5 个名称逐个添加。
5. 值只粘在这里，不要写进代码、README 或聊天。

如果找不到 `Settings`：确认你当前在仓库首页，不是 GitHub 个人主页；窗口窄时 `Settings` 可能收在右上角 `...` 菜单里。

## 四、开启自动运行

1. 点仓库上方 `Actions`。
2. 如果提示 workflow 未启用，点 `I understand my workflows, go ahead and enable them`。
3. 左侧选 `邮箱账单自动入库`。
4. 点 `Run workflow` → `Run workflow` 手动测试一次。
5. 看运行日志，出现 `本次共新增待入库 X 台` 就成功。

之后每 5 分钟自动跑一次。GitHub 的定时任务可能延迟几分钟到十几分钟，这是正常的。

## 五、把账单发到这个邮箱

到采货侠 / 拍机堂后台，把账单收件邮箱改成这个网易邮箱。

第一次收到账单后，建议：

1. 在网易邮箱里把采货侠 / 拍机堂的发件地址加入白名单，防止进垃圾箱。
2. 如果账单没被识别，把邮件主题和发件地址记下来，改 `BILL_KEYWORDS` 里的关键词。
3. 拍机堂如果没有附件、只有下载链接，需要再补一个链接下载逻辑。

## 账单口径

- 采货侠：成本优先取「实付金额」，没有就取「商品金额」；成色按「成色/等级」映射。
- 拍机堂：成本 = 绝对值「支出合计」+「其他调整明细」里的支付通道费、发货服务费分摊；同一订单多台按台分摊，尾差补到最后 1 台。
- 拍机堂单号：用「物品编号」做唯一单号，订单号写进备注，所以同一订单 7 台也能全部进待入库。
- 拍机堂账单没有成色字段，默认记「待定」，到货后在库存里改。

## 本地测试解析

```bash
pip install -r requirements.txt
python mail_to_inbox.py --file 采货侠账单.xlsx
python mail_to_inbox.py --file 2026年11月账单.xlsx
```

只会打印解析统计，不会写 Supabase。

## 安全提醒

- 授权码等同于邮箱密码，只放 GitHub Secrets。
- 建议这个网易邮箱只用来收账单，不要绑其他重要业务。
- 如果授权码泄露，去网易邮箱删除该授权码并重新生成。
