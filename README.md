# Emerald
⚠ 该项目现已永久停止维护，曾参与多次比赛。是我使用vb6开发的最后一个工程。
基于GDI双缓冲和GDI+，Bass的UI/游戏框架。  
  
Emerald框架可以应用于UI实现和游戏制作等场景。封装了windows底层绘图接口GDI和GDI+，利用双缓冲技术提高绘制效率，同时调用了第三方音频接口Bass。  

# 支持
Visual Basic 6.0  

# 文档
 [Github：Emerald Wiki](https://github.com/buger404/Emerald/wiki)  
 [码云：Emerald Wiki](https://gitee.com/buger404/Emerald/wikis)

# 制作初衷
不难发现，使用vb6开发出现代化的界面以及游戏都是比较困难的，对于新手则只能使用图片框控件，“堆控件”的形式实现游戏。Vb6不支持半透明图片不说，“堆控件”执行效率也低，而自绘游戏需要调用图形接口，这对新手而言也有不低的门槛。即使是有能力的人，直接调用图形接口也是很麻烦的，这又需要自己去造轮子。    
因此，我制作了一个便于开发的框架，集成了不少可能使用到的功能，降低开发门槛，方便开发者创造自己的世界。框架应用广泛，基本在各种需求都可以通用，这就不需要开发者为每个项目制作一个单独的轮子了。    
创作时，我借鉴了Unity的部分绘图机制，html + css，以及思考了一些能够更方便开发者开发的改进。例如，这个框架的元素不需要声明、实例化、继承这类操作，只需要在对应元素绘制的代码下方简单地调用一个不需要参数的函数，便可以得知：元素被点击了，按下的是哪个键，鼠标的坐标是多少等等。    
为了弥补不能“拖控件”可视化设计的问题，我给框架加入了“开发者模式”的功能，有便于开发者确定元素的坐标，节约时间。同时为开发者提供了整个页面控制器和主窗口的代码模板，用户可以通过模板快速创建页面。    
我自己也会使用自己的框架去开发一些软件，游戏，从中发现一些框架的不足，或是可以加入的方便开发者的功能，然后对框架加以改良。    
框架的功能是“组件化”的，这意味着只需要稍加修改，就可以将框架的某个功能单独分离出来，引入到别的项目中，这也是框架使用源代码形式的原因之一。开放源代码，同时还可以方便开发者自行修改，个性化框架使其更符合自己的需要。框架的构建也会因为开放源代码而更加高效，安全。“组件化”的另外一个好处是，当我需要重构框架的某个功能时，可以最小化重构带给其他功能的影响，提高开发效率。

# 第三方
* 第三方音频接口bass：http://www.un4seen.com/  
涉及的文件：Bass.dll，Bass.bas  
* vIstaswx GDI+ 声明模块  
涉及的文件：Gdiplus.bas  
* Win32API之vb6头文件made by 棉花糖  
涉及的文件：Win32Api.tlb  
