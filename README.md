# PPT2Image

PPT2Image is a library to Convert a PPT or PPTX file to Images by per slide.


## Usage

```

File file = new File("D:\\git\\PPT2Image\\1.pptx");
List<String> images = convertPPTtoImage(file,"D:\\git\\PPT2Image\\images\\pptx");

```

## Test

1. the quality of the images converted from PPTX is higher than those converted from PPT.
1. converting per slide from PPTX almost cost 1.6s in my computer, and PPT is 1s.

## Images

- PPTX 转换效果
![PPTX1](https://raw.githubusercontent.com/kkstudy/PPT2Image/master/images/pptx/1.jpg)
![PPTX2](https://raw.githubusercontent.com/kkstudy/PPT2Image/master/images/pptx/2.jpg)
![PPTX3](https://raw.githubusercontent.com/kkstudy/PPT2Image/master/images/pptx/3.jpg)


- PPT 转换效果
![PPT1](https://raw.githubusercontent.com/kkstudy/PPT2Image/master/images/ppt/1.jpg)
![PPT2](https://raw.githubusercontent.com/kkstudy/PPT2Image/master/images/ppt/2.jpg)
![PPT3](https://raw.githubusercontent.com/kkstudy/PPT2Image/master/images/ppt/3.jpg)
