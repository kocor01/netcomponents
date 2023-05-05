# 脚本使用教程

## python版本推荐
### python 3.8
  
## 安装脚本需要的包
### ```pip install requests```
### ```pip install xlrd```
### ```pip install xlwt```
### ```pip install bs4```

## 执行前准备 
### D盘下创建 `netcomponents` 目录
### 需要执行的文件放在 `netcomponents` 下，文件名称 NXP.xls
### 登录 www.netcomponents.com 从网站 cookie 里获取 login_auto 的值

## 执行脚本 
### 执行命令：```python .\netcomponents.py --login_auto="xxxxxxxx"```
### 执行完毕后，在 D盘 `netcomponents` 目录下生成 NXP_NEW 前缀的文件