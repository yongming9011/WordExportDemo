
# 一、前言

目前word导出方案主要有三种，各自对应的优缺点如下：

1. Apache POI导出。这个没的说，优点支持各种自定义操作，缺点就是代码特别多，又臭又长，而且不适合对导出格式要求特别高的需求，因为poi支持的格式比较少，而且设置格式有时候会出现莫名其妙的不生效的问题。
2. 第三方开源库easyPOI。这个库优点是简单，代码量少，而且支持word插入表格跟图片，只需要做一个docx的模板，并且将里面的变量替换成`{{}}`这种格式的，导出的文档完美保留格式。缺点是只支持docx格式，不支持在表格中循环插图片，制作模板的时候要特别小心，有时候一不小心多出一个空格都会导致数据无法显示，排查起来比较难，且该库的作者最近对于issue基本上不回复也不处理，bug还是比较多的。
3. 使用freemarker模板引擎导出。优点：导出速度快，而且支持各种表格，图片，表格循环插入图片，由于使用的是模板引擎，可以借助模板语法导出复杂格式的word，再者，由于使用的是模板，所以导出的格式不会出现问题，适合对格式要求高的导出。缺点：只支持doc格式。

# 二、freemarker用法

1. **创建模板**。使用office新建一个模板文件，最好是docx格式的。里面的变量内容最好用标识性明显的词语替代，不然保存成xml后不容易找到。而且创建模板的时候输入内容最好不要中英文混合，由于中英文的字体不一样，这样保存成xml一般中英文会分开显示，导致需要删除多余的内容，替换起来比较麻。如果有图片的话，最好放一张图片进去占位，生成的xml文件中就会显示一串base64字符串，到时候直接用变量替换就行，比如我创建了这样一个模板。

![img](https://i.loli.net/2020/06/24/3h47BWGtmyjQAk8.png)

2. **将模板另存为Word 2003 xml类型的xml**。一定要是Word 2003 xml，不能是Word xml，不然会无法打开。然后打开该xml文件，使用格式化工具进行格式化，推荐vscode（不过vscode会将一些类似`<w:t></w:t>`这种的直接格式化成`<w:t/>`这样的，会导致模板中的空白行失效，需要注意），然后将变量使用`${}`替换掉。

   ![image-20200624104857262.png](https://i.loli.net/2020/06/24/3zHd1SMPuJUaAfo.png)

   比如我需要生成一个循环表格，那么只需要在模板文件中将表格的第2行的所有内容用freemarker的遍历方法包裹起来就行如下：

   ![image-20200624105140212](https://i.loli.net/2020/06/24/kM9w3aTKnNPICQd.png)

   图片只需要将占位图生成的Base64字符串用变量替换掉即可，如下图，将这段Base64用变量替换掉：

   ![image-20200624105350265](https://i.loli.net/2020/06/24/2CDlRfZ3EFGLsqw.png)

   如果图片是在列表中循环生成的，则需要注意要将两个地方也要改成变量，不然生成的图片都跟第一张一样，一个是`<w:binData w:name`属性，另一个是	`<v:imagedata src`属性，原来的值是`<w:binData w:name="wordml://03000001.png"`修改为 `w:name="${"wordml://0300000"+person_index+".png"}"`，即在原有的属性基础上加上循环的索引，这样导出的图片就不会变成一样的了。

   ![image-20200624105712045](https://i.loli.net/2020/06/24/tpPrDeXAO48F7Qy.png)

   如果导出表格并不是每一列都有图片，而是有的有，有的没有，则需要对该单元格进行特殊处理，可以将`<w:pict></w:pict>`中的内容使用freemarker的`<#if></#if>`标签进行判断，如果有图片生成图片控件，不然如果没有base64字符串，会显示一个图片无法打开的奇怪内容。

3. **执行导出。**直接写一个工具类。传入模板的名称，导出的数据，导出的文件名，执行浏览器下载。

   ```java
   public class ExportUtils {
       public static void template2Word(String templateName, Map<String, Object> data, String fileName, 
                                        HttpServletResponse response) throws Exception{
           fileName = new String(fileName.getBytes("GBK"), StandardCharsets.ISO_8859_1) + ".doc";
           response.setContentType("application/x-msdownload;charset=UTF-8");
           response.addHeader("Content-Disposition", "attachment;filename=" + fileName + ";charset=UTF-8");
   
           Configuration configuration = new Configuration(Configuration.getVersion());
           configuration.setDefaultEncoding("UTF-8");
           // 设置模板存放的路径
           configuration.setClassForTemplateLoading(ExportUtils.class, "/template");
           Writer writer = new OutputStreamWriter(response.getOutputStream());
           Template t = configuration.getTemplate(templateName);
           t.process(data, writer);
       }
   }
   ```

# 三、FAQ
1. **特殊字符的处理：** 需要读取内容中，含有特殊字符，如：`< > @ ! $ &` 等等，可直接在模板中使用 `<![CDATA[  ]]>` 和 `?html` 处理。

   ```xml
   <w:t>${index}．<![CDATA[ ${quetion.questionTitle} ]]></w:t>
   ```
   或者
   
   ```xml
   <w:t>${info.name?html}</w:t>
   ```
