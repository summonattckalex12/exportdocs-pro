import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, BorderStyle, ShadingType, ImageRun, HorizontalPositionAlign, VerticalPositionAlign, HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom, TextWrappingType, Header, Footer, PageNumber, TabStopType, TabStopPosition, VerticalAlign, PageBreak, HeadingLevel, TableOfContents, LevelFormat, PageOrientation } from "docx";
import fs from "node:fs";

const bg = fs.readFileSync("/dev-server/src/assets/cover-bg.jpeg");
const mii = fs.readFileSync("/dev-server/src/assets/mii-logo.png");

const doc = new Document({
  creator: "test",
  styles: { default: { document: { run: { font: "Calibri", size: 22 } } } },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: {top:1440,right:1440,bottom:1440,left:1440} }, titlePage: true },
    headers: {
      first: new Header({ children:[new Paragraph({children:[new TextRun("")]})]}),
      default: new Header({ children:[
        new Table({ width:{size:9026,type:WidthType.DXA}, columnWidths:[1400,6226,1400], rows:[new TableRow({children:[
          new TableCell({borders:{top:{style:BorderStyle.NONE,size:0,color:"FFFFFF"},bottom:{style:BorderStyle.NONE,size:0,color:"FFFFFF"},left:{style:BorderStyle.NONE,size:0,color:"FFFFFF"},right:{style:BorderStyle.NONE,size:0,color:"FFFFFF"}}, width:{size:1400,type:WidthType.DXA}, children:[new Paragraph({children:[new ImageRun({type:"png",data:mii,transformation:{width:40,height:40},altText:{title:"l",description:"l",name:"l"}})]})]}),
          new TableCell({borders:{top:{style:BorderStyle.NONE,size:0,color:"FFFFFF"},bottom:{style:BorderStyle.NONE,size:0,color:"FFFFFF"},left:{style:BorderStyle.NONE,size:0,color:"FFFFFF"},right:{style:BorderStyle.NONE,size:0,color:"FFFFFF"}}, width:{size:6226,type:WidthType.DXA}, children:[new Paragraph({children:[new TextRun("hdr")]})]}),
          new TableCell({borders:{top:{style:BorderStyle.NONE,size:0,color:"FFFFFF"},bottom:{style:BorderStyle.NONE,size:0,color:"FFFFFF"},left:{style:BorderStyle.NONE,size:0,color:"FFFFFF"},right:{style:BorderStyle.NONE,size:0,color:"FFFFFF"}}, width:{size:1400,type:WidthType.DXA}, children:[new Paragraph({children:[new ImageRun({type:"png",data:mii,transformation:{width:40,height:40},altText:{title:"r",description:"r",name:"r"}})]})]}),
        ]})]})
      ]})
    },
    footers: { first: new Footer({children:[new Paragraph({children:[new TextRun("")]})]}), default: new Footer({children:[new Paragraph({children:[new TextRun("f")]})]}) },
    children: [
      new Paragraph({children:[new ImageRun({type:"jpg",data:bg,transformation:{width:794,height:1123},altText:{title:"bg",description:"bg",name:"bg"},floating:{horizontalPosition:{relative:HorizontalPositionRelativeFrom.PAGE,align:HorizontalPositionAlign.LEFT},verticalPosition:{relative:VerticalPositionRelativeFrom.PAGE,align:VerticalPositionAlign.TOP},behindDocument:true,zIndex:0,allowOverlap:true,wrap:{type:TextWrappingType.NONE}}})]}),
      new Paragraph({children:[new PageBreak()]}),
      new Paragraph({heading:HeadingLevel.HEADING_1, children:[new TextRun("Section 2")]}),
      new Paragraph({children:[new TextRun("body")]}),
    ],
  }],
});
const buf = await Packer.toBuffer(doc);
fs.writeFileSync("/tmp/docxtest/out.docx", buf);
console.log("wrote", buf.length);
