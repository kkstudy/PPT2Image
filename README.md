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

