## The problem

I am trying to determine the page on which a partciular style resides.  Some documents work while others don't generate the correct result.

The code snippet takes a docx file and increments the page variable whenever it encounters on fo the following:

- **<w:lastRenderedPageBreak//>**
- **<w:br w:type="page"//>**
- **<w:sectPr** not having a **<w:type w:val="continuous"/>**

I encountered documents in which two of the tags were in sequence in two seperate **<w:p **.  These had to be counted as one.

The solution generats a csv file that can be opened in a spreadsheet program.

## WorksOK

The document **WorksOK.docx** works OK. The output of the csv is below:
| Style           | Style Text                       | Page |
| --------------- | -------------------------------- | ---- |
| RACIResp        | Opportunity (Page 1 / Section 1) | 1    |
| RACIAccountable | Weakness (Page 1 / Section 1)    | 1    |
| RACIAccountable | Weakness (Page 2  / Section 1)   | 2    |
| RACIInf         | Strengths (Page 2 / Section 2)   | 2    |
| RACIAccountable | Weakness (Page 5 / Section 3)    | 5    |
| RACIInf         | Strength (Page 6 / Section 4)    | 6    |

## Problem

The document **Problem.docx** is not able to find a page boundary marker between 2 and 3. As a result the output is off.

| Style           | Style Text        | Page |
| --------------- | ----------------- | ---- |
| RACIResp        | Responsible: Pg 1 | 1    |
| RACIInf         | Informed: Pg 1    | 1    |
| RACIResp        | Responsible: Pg3  | 2    |
| RACIAccountable | Accountable: Pg3  | 2    |

### the XML of **Problem.docx** is hereunder.  I could not find the page break using a visual inspection. The only one I found is `<w:p w14:paraId="45B2C77B"`

## Solution

If you have a solution please drop a note. Thanks

```
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
    xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
    xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"
    xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex"
    xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex"
    xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex"
    xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex"
    xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex"
    xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex"
    xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink"
    xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d"
    xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:oel="http://schemas.microsoft.com/office/2019/extlst"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
    xmlns:v="urn:schemas-microsoft-com:vml"
    xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:w10="urn:schemas-microsoft-com:office:word"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
    xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"
    xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
    xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml"
    xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"
    xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"
    xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
    xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
    xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
    xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14">
    <w:body>
        <w:p w14:paraId="73B496DE" w14:textId="6A7DB9C7" w:rsidR="00166C05" w:rsidRPr="00F44110" w:rsidRDefault="589F8884" w:rsidP="2A6F2D8D">
            <w:pPr>
                <w:pStyle w:val="Heading1"/>
                <w:rPr>
                    <w:rStyle w:val="Heading1Char"/>
                    <w:b/>
                </w:rPr>
            </w:pPr>
            <w:bookmarkStart w:id="1" w:name="_2_-_Create"/>
            <w:bookmarkStart w:id="2" w:name="_Toc158390227"/>
            <w:bookmarkEnd w:id="1"/>
            <w:r w:rsidRPr="7CEC6C9D">
                <w:rPr>
                    <w:rStyle w:val="Heading1Char"/>
                    <w:b/>
                </w:rPr>
                <w:t>2</w:t>
            </w:r>
            <w:r w:rsidR="00C4559A">
                <w:rPr>
                    <w:rStyle w:val="Heading1Char"/>
                    <w:b/>
                </w:rPr>
                <w:t xml:space="preserve"></w:t>
            </w:r>
            <w:r w:rsidR="00974F7F">
                <w:rPr>
                    <w:rStyle w:val="Heading1Char"/>
                    <w:b/>
                </w:rPr>
                <w:t>–</w:t>
            </w:r>
            <w:r w:rsidRPr="7CEC6C9D">
                <w:rPr>
                    <w:rStyle w:val="Heading1Char"/>
                    <w:b/>
                </w:rPr>
                <w:t xml:space="preserve"></w:t>
            </w:r>
            <w:bookmarkEnd w:id="2"/>
            <w:r w:rsidR="00974F7F">
                <w:rPr>
                    <w:rStyle w:val="Heading1Char"/>
                    <w:b/>
                </w:rPr>
                <w:t>Pg 1</w:t>
            </w:r>
            <w:r w:rsidRPr="7CEC6C9D">
                <w:rPr>
                    <w:rStyle w:val="Heading1Char"/>
                    <w:b/>
                </w:rPr>
                <w:t xml:space="preserve"></w:t>
            </w:r>
        </w:p>
        <w:p w14:paraId="45C61C11" w14:textId="6D72A275" w:rsidR="6A3B3D05" w:rsidRDefault="6A3B3D05" w:rsidP="002630A8">
            <w:pPr>
                <w:pStyle w:val="RACIResp"/>
                <w:rPr>
                    <w:highlight w:val="lightGray"/>
                </w:rPr>
            </w:pPr>
            <w:r w:rsidRPr="002630A8">
                <w:rPr>
                    <w:highlight w:val="lightGray"/>
                </w:rPr>
                <w:t xml:space="preserve">Responsible: </w:t>
            </w:r>
            <w:r w:rsidR="00974F7F">
                <w:rPr>
                    <w:highlight w:val="lightGray"/>
                </w:rPr>
                <w:t>Pg 1</w:t>
            </w:r>
        </w:p>
        <w:p w14:paraId="5B067CF0" w14:textId="77777777" w:rsidR="00335C08" w:rsidRDefault="00335C08" w:rsidP="00335C08">
            <w:pPr>
                <w:rPr>
                    <w:highlight w:val="lightGray"/>
                </w:rPr>
            </w:pPr>
        </w:p>
        <w:p w14:paraId="026750B7" w14:textId="77777777" w:rsidR="00335C08" w:rsidRDefault="00335C08" w:rsidP="00335C08">
            <w:pPr>
                <w:rPr>
                    <w:highlight w:val="lightGray"/>
                </w:rPr>
            </w:pPr>
        </w:p>
        <w:p w14:paraId="48418805" w14:textId="77777777" w:rsidR="00335C08" w:rsidRDefault="00335C08" w:rsidP="00335C08">
            <w:pPr>
                <w:rPr>
                    <w:highlight w:val="lightGray"/>
                </w:rPr>
            </w:pPr>
        </w:p>
        <w:p w14:paraId="0A26AA82" w14:textId="77777777" w:rsidR="00335C08" w:rsidRDefault="00335C08" w:rsidP="00335C08">
            <w:pPr>
                <w:rPr>
                    <w:highlight w:val="lightGray"/>
                </w:rPr>
            </w:pPr>
        </w:p>
        <w:p w14:paraId="14D22EDA" w14:textId="77777777" w:rsidR="00335C08" w:rsidRDefault="00335C08" w:rsidP="00335C08">
            <w:pPr>
                <w:rPr>
                    <w:highlight w:val="lightGray"/>
                </w:rPr>
            </w:pPr>
        </w:p>
        <w:p w14:paraId="1001A019" w14:textId="77777777" w:rsidR="00335C08" w:rsidRDefault="00335C08" w:rsidP="00335C08">
            <w:pPr>
                <w:rPr>
                    <w:highlight w:val="lightGray"/>
                </w:rPr>
            </w:pPr>
        </w:p>
        <w:p w14:paraId="0393FEF7" w14:textId="77777777" w:rsidR="00335C08" w:rsidRDefault="00335C08" w:rsidP="00335C08">
            <w:pPr>
                <w:rPr>
                    <w:highlight w:val="lightGray"/>
                </w:rPr>
            </w:pPr>
        </w:p>
        <w:p w14:paraId="6F601EEF" w14:textId="77777777" w:rsidR="00335C08" w:rsidRDefault="00335C08" w:rsidP="00335C08">
            <w:pPr>
                <w:rPr>
                    <w:highlight w:val="lightGray"/>
                </w:rPr>
            </w:pPr>
        </w:p>
        <w:p w14:paraId="042887FC" w14:textId="77777777" w:rsidR="00335C08" w:rsidRDefault="00335C08" w:rsidP="00335C08">
            <w:pPr>
                <w:rPr>
                    <w:highlight w:val="lightGray"/>
                </w:rPr>
            </w:pPr>
        </w:p>
        <w:p w14:paraId="5F9BB39B" w14:textId="2EC3E47B" w:rsidR="00335C08" w:rsidRPr="00335C08" w:rsidRDefault="00335C08" w:rsidP="00335C08">
            <w:pPr>
                <w:pStyle w:val="RACIInf"/>
                <w:rPr>
                    <w:highlight w:val="lightGray"/>
                </w:rPr>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:highlight w:val="lightGray"/>
                </w:rPr>
                <w:t>Informed: Pg 1</w:t>
            </w:r>
        </w:p>
        <w:p w14:paraId="5633FD1D" w14:textId="1C7E9C51" w:rsidR="00166C05" w:rsidRDefault="00166C05" w:rsidP="00166C05">
            <w:pPr>
                <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
                <w:rPr>
                    <w:rFonts w:eastAsia="Times New Roman"/>
                </w:rPr>
            </w:pPr>
        </w:p>
        <w:p w14:paraId="53103026" w14:textId="4D9A3CDB" w:rsidR="009B26F1" w:rsidRDefault="009B26F1" w:rsidP="00166C05">
            <w:pPr>
                <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
                <w:rPr>
                    <w:rFonts w:eastAsia="Times New Roman"/>
                </w:rPr>
            </w:pPr>
        </w:p>
        <w:p w14:paraId="2C5C4017" w14:textId="59FA4550" w:rsidR="007660D7" w:rsidRDefault="007660D7">
            <w:pPr>
                <w:rPr>
                    <w:rFonts w:eastAsia="Times New Roman"/>
                </w:rPr>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:rFonts w:eastAsia="Times New Roman"/>
                </w:rPr>
                <w:br w:type="page"/>
            </w:r>
        </w:p>
        <w:p w14:paraId="7141F576" w14:textId="77777777" w:rsidR="007660D7" w:rsidRDefault="007660D7" w:rsidP="00166C05">
            <w:pPr>
                <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
                <w:rPr>
                    <w:rFonts w:eastAsia="Times New Roman"/>
                </w:rPr>
            </w:pPr>
        </w:p>
        <w:p w14:paraId="2F13F6DA" w14:textId="6E1DC1CB" w:rsidR="00166C05" w:rsidRDefault="00166C05" w:rsidP="00166C05">
            <w:pPr>
                <w:pStyle w:val="Heading2"/>
                <w:rPr>
                    <w:rFonts w:eastAsia="Times New Roman"/>
                </w:rPr>
            </w:pPr>
            <w:r w:rsidRPr="2A6F2D8D">
                <w:rPr>
                    <w:rFonts w:eastAsia="Times New Roman"/>
                </w:rPr>
                <w:t xml:space="preserve">I </w:t>
            </w:r>
            <w:r w:rsidR="00B26BD1">
                <w:rPr>
                    <w:rFonts w:eastAsia="Times New Roman"/>
                </w:rPr>
                <w:t>Page 2</w:t>
            </w:r>
        </w:p>
        <w:p w14:paraId="2BF1C2E4" w14:textId="77777777" w:rsidR="00166C05" w:rsidRDefault="00166C05" w:rsidP="00166C05">
            <w:pPr>
                <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
                <w:rPr>
                    <w:rFonts w:eastAsia="Times New Roman"/>
                </w:rPr>
            </w:pPr>
        </w:p>
        <w:p w14:paraId="301500D9" w14:textId="77777777" w:rsidR="00166C05" w:rsidRDefault="00166C05" w:rsidP="00166C05">
            <w:pPr>
                <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
                <w:rPr>
                    <w:rFonts w:eastAsia="Times New Roman"/>
                </w:rPr>
            </w:pPr>
        </w:p>
        <w:p w14:paraId="30CE4BCA" w14:textId="29F565C5" w:rsidR="00166C05" w:rsidRDefault="00166C05" w:rsidP="00166C05">
            <w:pPr>
                <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
                <w:rPr>
                    <w:rFonts w:eastAsia="Times New Roman"/>
                </w:rPr>
            </w:pPr>
        </w:p>
        <w:p w14:paraId="2D916A6D" w14:textId="77777777" w:rsidR="00360A06" w:rsidRDefault="00360A06" w:rsidP="00360A06"/>
        <w:p w14:paraId="1293F478" w14:textId="1384FC3C" w:rsidR="00166C05" w:rsidRDefault="00166C05" w:rsidP="00360A06"/>
        <w:p w14:paraId="45B2C77B" w14:textId="77777777" w:rsidR="00F9091B" w:rsidRDefault="00F9091B" w:rsidP="00360A06">
            <w:pPr>
                <w:sectPr w:rsidR="00F9091B" w:rsidSect="00A74BE7">
                    <w:headerReference w:type="default" r:id="rId7"/>
                    <w:footerReference w:type="default" r:id="rId8"/>
                    <w:pgSz w:w="11906" w:h="16838"/>
                    <w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>
                    <w:cols w:space="708"/>
                    <w:docGrid w:linePitch="360"/>
                </w:sectPr>
            </w:pPr>
        </w:p>
        <w:p w14:paraId="7A215D5B" w14:textId="16822E3D" w:rsidR="00310D9A" w:rsidRPr="00F44110" w:rsidRDefault="00DB4497" w:rsidP="2A6F2D8D">
            <w:pPr>
                <w:pStyle w:val="Heading1"/>
                <w:rPr>
                    <w:rStyle w:val="Heading1Char"/>
                    <w:b/>
                </w:rPr>
            </w:pPr>
            <w:bookmarkStart w:id="3" w:name="_Toc158390228"/>
            <w:r>
                <w:rPr>
                    <w:rStyle w:val="Heading1Char"/>
                    <w:b/>
                </w:rPr>
                <w:lastRenderedPageBreak/>
                <w:t>3</w:t>
            </w:r>
            <w:r w:rsidR="4011B11E" w:rsidRPr="7CEC6C9D">
                <w:rPr>
                    <w:rStyle w:val="Heading1Char"/>
                    <w:b/>
                </w:rPr>
                <w:t xml:space="preserve"></w:t>
            </w:r>
            <w:r w:rsidR="00974F7F">
                <w:rPr>
                    <w:rStyle w:val="Heading1Char"/>
                    <w:b/>
                </w:rPr>
                <w:t>–</w:t>
            </w:r>
            <w:r w:rsidR="73A2F062" w:rsidRPr="7CEC6C9D">
                <w:rPr>
                    <w:rStyle w:val="Heading1Char"/>
                    <w:b/>
                </w:rPr>
                <w:t xml:space="preserve"></w:t>
            </w:r>
            <w:bookmarkEnd w:id="3"/>
            <w:r w:rsidR="00974F7F">
                <w:rPr>
                    <w:rStyle w:val="Heading1Char"/>
                    <w:b/>
                </w:rPr>
                <w:t>Pg 3</w:t>
            </w:r>
        </w:p>
        <w:p w14:paraId="78455D8F" w14:textId="0DA89A35" w:rsidR="008D7069" w:rsidRDefault="00270EEE" w:rsidP="0099448C">
            <w:pPr>
                <w:pStyle w:val="RACIResp"/>
            </w:pPr>
            <w:r>
                <w:t>Responsible:</w:t>
            </w:r>
            <w:r w:rsidR="00974F7F">
                <w:t xml:space="preserve"> Pg3</w:t>
            </w:r>
        </w:p>
        <w:p w14:paraId="3ABF2C6F" w14:textId="044D7104" w:rsidR="00BB4613" w:rsidRDefault="006B32B2" w:rsidP="004F3842">
            <w:pPr>
                <w:pStyle w:val="RACIAccountable"/>
                <w:rPr>
                    <w:rStyle w:val="RACIAccountableChar"/>
                </w:rPr>
            </w:pPr>
            <w:r w:rsidRPr="001D597A">
                <w:rPr>
                    <w:rStyle w:val="RACIAccountableChar"/>
                </w:rPr>
                <w:t xml:space="preserve">Accountable: </w:t>
            </w:r>
            <w:r w:rsidR="00974F7F">
                <w:rPr>
                    <w:rStyle w:val="RACIAccountableChar"/>
                </w:rPr>
                <w:t>Pg3</w:t>
            </w:r>
        </w:p>
        <w:p w14:paraId="058FEE15" w14:textId="0F16770E" w:rsidR="00335C08" w:rsidRPr="00335C08" w:rsidRDefault="00335C08" w:rsidP="00335C08"/>
        <w:p w14:paraId="158B8A79" w14:textId="77777777" w:rsidR="00685162" w:rsidRDefault="00685162" w:rsidP="7CEC6C9D">
            <w:pPr>
                <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
            </w:pPr>
        </w:p>
        <w:sectPr w:rsidR="00685162" w:rsidSect="00A74BE7">
            <w:headerReference w:type="default" r:id="rId9"/>
            <w:footerReference w:type="default" r:id="rId10"/>
            <w:pgSz w:w="11906" w:h="16838"/>
            <w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>
            <w:cols w:space="708"/>
            <w:docGrid w:linePitch="360"/>
        </w:sectPr>
    </w:body>
</w:document>
```
