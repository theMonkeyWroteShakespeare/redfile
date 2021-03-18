
from .downfile import downfile
from .until import *
class xinhan(downfile):
    def __init__(self,data):
        downfile.__init__(self,data)

    def setPage(self):  # 页面页面字号设置
        # 页面设置
        # 国家公文格式标准要求是上边距版心3.7cm
        # 但是如果简单的把上边距设置为3.7cm
        # 则因为文本的第一行本身有行距
        # 会导致实际版心离上边缘较远，上下边距设置为3.3cm
        # 是经过实验的，可以看看公文标准的图示
        # 版心指的是文字与边缘距离
        self.doc.PageSetup.TopMargin = 3.3 * cm_to_points
        # 上边距3.3厘米
        self.doc.PageSetup.BottomMargin = 2.0 * cm_to_points
        # 下边距3.3厘米
        self.doc.PageSetup.LeftMargin = 2.8 * cm_to_points
        # 左边距2.8厘米
        self.doc.PageSetup.RightMargin = 2.6 * cm_to_points
        # 右边距2.6厘米

        # 设置正常样式的字体
        # 是为了后面指定行和字符网格时
        # 按照这个字体标准进行
        self.doc.Styles(-1).Font.Name = '仿宋'
        # word中的“正常”样式字体为仿宋
        self.doc.Styles(-1).Font.NameFarEast = '仿宋'
        # word中的“正常”样式字体为仿宋
        self.doc.Styles(-1).Font.NameAscii = '仿宋'
        # word中的“正常”样式字体为仿宋
        self.doc.Styles(-1).Font.NameOther = '仿宋'
        # word中的“正常”样式字体为仿宋
        self.doc.Styles(-1).Font.Size = 16
        # word中的“正常”样式字号为三号

        self.doc.PageSetup.LayoutMode = 1
        # 指定行和字符网格
        self.doc.PageSetup.CharsLine = 28
        # 每行28个字
        self.doc.PageSetup.LinesPage = 23
        # 每页22行，会自动设置行间距

        # 页码设置
        self.doc.PageSetup.FooterDistance = 2.4 * cm_to_points
        # 页码距下边缘2.8厘米
        self.doc.PageSetup.DifferentFirstPageHeaderFooter = 0
        # 首页页码相同
        self.doc.PageSetup.OddAndEvenPagesHeaderFooter = 0
        # 页脚奇偶页相同
        w = self.doc.windows(1)
        # 获得文档的第一个窗口
        w.view.seekview = 4
        # 获得页眉页脚视图
        se = w.Selection
        # 获取窗口的选择对象
        se.headerfooter.pagenumbers.startingnumber = 1
        # 设置起始页码
        se.headerfooter.pagenumbers.NumberStyle = 0
        # 设置页码样式为单纯的阿拉伯数字
        se.WholeStory()
        # 扩选到整个部分（会选中整个页眉页脚）
        se.Delete()
        # 按下删除键，这两句是为了清除原来的页码
        se.headerfooter.pagenumbers.Add(4)
        # 添加页面外侧页码
        se.MoveLeft(1, 2)
        # 移动到页码左边，移动了两个字符距离
        se.TypeText('— ')
        # 给页码左边加上一字线，注意不是减号
        se.MoveRight()
        # 移动到页码末尾，移动了一个字符距离
        # 默认参数是1（字符）
        se.TypeText(' —')
        se.WholeStory()
        # 扩选到整个页眉页脚部分，此处是必要的
        # 否则s只是在输入一字线后的一个光标，没有选择区域
        se.Font.Name = '宋体'
        se.Font.Size = 14
        # 页码字号为四号
        se.paragraphformat.rightindent = 21
        # 页码向左缩进1字符（21磅）
        se.paragraphformat.leftindent = 21
        # 页码向右缩进1字符（21磅）
        self.doc.Styles('页眉').ParagraphFormat.Borders(-3).LineStyle = 0
        # 页眉无底边框横线

        w.view.seekview = 0
    def add_xin_han_title(self):  # 绘制信函格式的大红头和上边下边的双红线
        str=self.data["发文机关"]
        isredpaper=self.data["是否使用红头纸"]
        self.s.TypeText("\n")
        shape = self.doc.Shapes.AddTextbox(1, 79.38, 3 * cm_to_points - 8.93, 442.26, 80,
                                      self.doc.Range(0, 0))  # 使用文本框来写大红头，大红头字距离上边3cm,要剪掉字上边的空白
        shape.Line.Visible = 0
        shape.RelativeHorizontalPosition = 1
        shape.RelativeVerticalPosition = 1
        textbox = shape.TextFrame
        textbox.MarginBottom = 0
        textbox.MarginTop = 0
        textbox.MarginRight = 0
        textbox.MarginLeft = 0
        textbox.HorizontalAnchor = 2
        textbox.TextRange.Text = str[0]
        textbox.TextRange.font.Name = "方正小标宋简体"
        textbox.TextRange.font.Size = font_size["小初"]
        textbox.TextRange.font.Color = 255
        textbox.TextRange.ParagraphFormat.Alignment = 1  # 1是居中0 是靠左 2是靠右
        maxlen = len(str[0])  # 找出字数最长的单位
        if maxlen > 12:
            textbox.TextRange.Font.Scaling = int(12 * 100 / maxlen)
        self.drawTheRedLine(139.0645, self.doc.Range(0, 0))
        self.drawTheRedLine(143.0645, self.doc.Range(0, 0), 1.5)
        self.drawTheRedLine(29.7 * cm_to_points - 2 * cm_to_points - 5, self.doc.Range(0, 0), 1.5)
        self.drawTheRedLine(29.7 * cm_to_points - 2 * cm_to_points - 1, self.doc.Range(0, 0))

        self.s.TypeText("\n")
        self.s.Font.Scaling = 100

