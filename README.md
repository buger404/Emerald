# Emerald
面向游戏的轻量级绘图框架，仅限Visual Basic 6.0。

# 支持
最低：Visual Basic 6.0  
最高：Visual Basic 6.0 (SP 6)

# 组成
# Page Manager
管理```Pages```以及```过场特效```  
# 创建
```VBS
Dim Manager As GMan  
...
Set Manager = New GMan
```
# 页面的设置
```VBS
Manager.ActivePage = "MyPage"
```
# 过场特效
```VBS
NewTransform [Transform ID],[During]  
```
Transform ID  
- 0: Fade In
- 1: Fade Out  
```VBS
Manager.NewTransform 0,300
```

# Page
你的游戏页面，通过`Page Manager`进行管理。
# 创建
```VBS
Dim Page As GPage
Public Sub Update()
    ...
End Sub
...
Private Sub Class_Initialize()
    Set Page = New GPage
    Page.Create Me
    ...
    ECore.Add Page, "MyPage" '加入到Page Manager
End Sub
```
# 加入资源组
从指定文件夹中加载所有图片文件
```VBS
Page.NewImages [path]
```
