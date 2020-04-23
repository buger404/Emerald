Attribute VB_Name = "AboutMe"
'========================================================
'   Emerald
'   制作：Error 404（QQ 1361778219，陈志琰）
'   邮箱：ris_vb@126.com
'========================================================
'   面向VB6的轻量级绘图框架
'========================================================
'   组成（标*号的表示非作者编写）：
'   ┗━━━━调试功能
'       ┗━━━━Debuginfo.frm：显示调试详细信息使用的窗口
'       ┗━━━━DebugWindow.frm：显示调试工具栏使用的窗口
'       ┗━━━━GDebug.cls：调试工具栏的界面绘制
'       ┗━━━━DebugSwitch.bas：存放调试变量的模块
'   ┗━━━━*BASS（http://www.un4seen.com/）
'       (BASS is an audio library for use in software on several platforms.)
'       ┗━━━━*Bass.bas：Bass API 声明模块
'       ┗━━━━GMusic.cls：Emerald封装的使用了Bass的音乐播放器
'       ┗━━━━GMusicList.cls：音乐列表，管理GMusic
'   ┗━━━━存档
'       ┗━━━━GSaving.cls：Emerald存档功能
'       ┗━━━━BMEA_Engine.bas：黑嘴加密算法（不可逆）
'   ┗━━━━绘制
'       ┗━━━━GMan.cls：页面管理器，游戏核心，支持页面过渡
'       ┗━━━━GPage.cls：游戏页面，包含常规图形的绘制
'       ┗━━━━GFont.cls：游戏字体绘制
'       ┗━━━━GResource.cls：游戏资源管理
'   ┗━━━━动画
'       ┗━━━━GAnimation.cls：Emerald常规动画函数
'       ┗━━━━Animations.bas：Emerald图像动画模块
'   ┗━━━━碰撞箱
'       ┗━━━━GCrashBox.cls：碰撞箱功能
'   ┗━━━━*GDIPlus
'       ┗━━━━*Gdiplus.bas：vIstaswx GDI+ 声明模块
'   ┗━━━━其他
'       ┗━━━━AeroEffect.bas：Aero效果
'       ┗━━━━GCore.bas：Emerald部分核心功能
'       ┗━━━━GSysPage.cls：Emerald页面的绘制
'       ┗━━━━EmeraldWindow.frm：显示Emerald信息使用的窗口
'========================================================
