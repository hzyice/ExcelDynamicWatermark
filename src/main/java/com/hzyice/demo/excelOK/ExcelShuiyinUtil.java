package com.hzyice.demo.excelOK;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import java.util.ArrayList;
import java.util.List;


/*生成水印工具类*/

public class ExcelShuiyinUtil {
    private static ActiveXComponent xl = null; //Excel对象(防止打开多个)
    private static Dispatch workbooks = null;  //工作簿对象  
    private Dispatch workbook = null; //具体工作簿  
    private Dispatch sheets = null;// 获得sheets集合对象  
    private Dispatch currentSheet = null;// 当前sheet  
    /**
     * 打开excel文件 
     * @param filepath 文件路径名称 
     * @param visible  是否显示打开 
     * @param //readonly 是否只读方式打开
     */
    public  void OpenExcel(String filepath, boolean visible) {  
        try {  
            initComponents(); //清空原始变量  
            ComThread.InitSTA();  
            if(xl==null)  
                xl = new ActiveXComponent("Excel.Application"); //Excel对象  
            xl.setProperty("Visible", new Variant(visible));//设置是否显示打开excel  
            if(workbooks==null)  
                workbooks = xl.getProperty("Workbooks").toDispatch(); //打开具体工作簿  
                workbook = Dispatch.invoke(workbooks, "Open", Dispatch.Method,  
               new Object[] { filepath,  
                                        new Variant(true), // 是否以只读方式打开  
                                        new Variant(false),  
                                         "1",  
                                        "pwd" },   //输入密码"pwd",若有密码则进行匹配，无则直接打开  
                                         new int[1]).toDispatch();  
        } catch (Exception e) {  
            e.printStackTrace();  
            releaseSource();  
        }  
    }  
    /**
     * 工作簿另存为 
     * @param filePath 另存为的路径 
     * 例如 SaveAs="D:TEST/c.xlsx" 
     */
    private void SaveAs(String filePath){  
           Dispatch.call(workbook, "SaveAs",filePath);  
      }  
    /**
     * 关闭excel文档 
     * @param f 含义不明 （关闭是否保存？默认false） 
     */
    public  void CloseExcel(boolean f) {  
        try {  
            Dispatch.call(workbook, "Save");  
            Dispatch.call(workbook, "Close", new Variant(f));  
        } catch (Exception e) {  
            e.printStackTrace();  
        } finally {  
                releaseSource();  
        }    
    } 
    /**
     * 另存为文档并关闭excel文档 
     * @param //f 含义不明 （关闭是否保存？默认false）
     */
    public  void CloseExcel(String filePath) {  
        try {  
        	 Dispatch.call(workbook, "SaveAs",filePath);   
            Dispatch.call(workbook, "Close", new Variant(true));  
        } catch (Exception e) {  
            e.printStackTrace();  
        } finally {  
                releaseSource();  
        }    
    }
    /*
     * 初始化 
     * */
    private void initComponents(){  
        workbook = null;  
        currentSheet = null;  
        sheets = null;  
    }  
    /**
     * 释放资源 
     */
    private static void releaseSource(){  
        if(xl!=null){  
            xl.invoke("Quit", new Variant[] {});  
            xl = null;  
        }  
        workbooks = null;  
        ComThread.Release();  
        System.gc();  
    }  
    /**
     * 得到当前sheet 
     * @return 
     */
    private Dispatch getCurrentSheet() {  
        currentSheet = Dispatch.get(workbook, "ActiveSheet").toDispatch();  
        return currentSheet;  
    }  
    /**
     * 修改当前工作表的名字 
     * @param newName 
     */
    private void modifyCurrentSheetName(String newName) {  
        Dispatch.put(getCurrentSheet(), "name", newName);    
    }  
  
    /**
     * 得到当前工作表的名字 
     * @return 
     */
    private String getCurrentSheetName(Dispatch sheets) {  
        return Dispatch.get(sheets, "name").toString();  
    }  
    /**
     * 通过工作表名字得到工作表 
     * @param name sheetName 
     * @return 
     */
    private Dispatch getSheetByName(String name) {  
        return Dispatch.invoke(getSheets(), "Item", Dispatch.Get, new Object[]{name}, new int[1]).toDispatch();  
    }  
    /**
     *  得到sheets的集合对象 
     * @return 
     */
    private Dispatch getSheets() {  
        if(sheets==null)  
            sheets = Dispatch.get(workbook, "sheets").toDispatch();  
        return sheets;  
    }  
    /**
     * 通过工作表索引得到工作表(第一个工作簿index为1) 
     * @param index 
     * @return  sheet对象 
     */
    private Dispatch getSheetByIndex(Integer index) {  
        return Dispatch.invoke(getSheets(), "Item", Dispatch.Get, new Object[]{index}, new int[1]).toDispatch();  
    }  
  
    /**
     * 得到sheet的总数 
     * @return 
     */
    private int getSheetCount() {  
        int count = Dispatch.get(getSheets(), "count").toInt();  
        return count;  
    }  
    /**
     * 给所有的sheet添加背景 
     * @param filepath 图片路径 
     */
    public void setBlackGroudPrituce(String filepath)  
    {  
        int num=this.getSheetCount();  
        for (int i = 1; i <= num; i++) {  
            Dispatch sheets=this.getSheetByIndex(i);  
         
            Dispatch.call(sheets,"SetBackgroundPicture",filepath);
            
    
            
        }     
    }  
    /**
     *  添加新的工作表(sheet)，并且隐藏（添加后为默认为当前激活的工作表） 
     */
    public void addSheet(String name) {  
//      for (int i = 1; i <= this.getSheetCount(); i++) {  
//          Dispatch sheets=this.getSheetByIndex(i);  
//         if(name.equals(this.getCurrentSheetName(sheets)))  
//            {  
//                return false;  
//             }             
//         }   
          currentSheet=Dispatch.get(Dispatch.get(workbook, "sheets").toDispatch(), "add").toDispatch();  
        //  Dispatch.put(currentSheet,"Name",name);  
          Dispatch.put(currentSheet, "Visible", new Boolean(false));  
          System.out.println("插入信息为:"+name);  
    }  
    /**
     * 得到工作薄的名字 
     * @return 
     */
    private String getWorkbookName() {  
        if(workbook==null)  
            return null;  
        return Dispatch.get(workbook, "name").toString();  
    }  
    /**
     *  获取所有表名 
     */
    public List findSheetName()  
    {  
        int num=this.getSheetCount();  
        List list=new ArrayList();  
        for (int i = 1; i <= num; i++) {  
            currentSheet=this.getSheetByIndex(i);  
            list.add(this.getCurrentSheetName(currentSheet));    
           }   
        return list;  
    }  
    /**
     * 设置页脚信息 
     */
    private void setFooter(String foot) {  
        currentSheet=this.getCurrentSheet();  
        Dispatch PageSetup=Dispatch.get(currentSheet,"PageSetup").toDispatch();  
        Dispatch.put(PageSetup,"CenterFooter",foot);  
    }  
    /**
     * 获取页脚信息 
     */
    private String getFooter() {  
        currentSheet=this.getCurrentSheet();  
        Dispatch PageSetup=Dispatch.get(currentSheet,"PageSetup").toDispatch();  
        return Dispatch.get(PageSetup,"CenterFooter").toString();  
    }  
    /**
     * 锁定工作簿 
     */
    public void setPassword() {  
        Dispatch.call(workbook, "Protect",123,true,false);  
    }  
    /**
     * 设置名称管理器 
     * @param name 名称管理器名 不能以数字或者下划线开头，中间不能包含空格和其他无效字符 
     * @param comment 备注 
     * @param place 备注位置  
     * @return 
     */
    public void setName(String name,String place,String comment) {  
        Dispatch Names=Dispatch.get(workbook, "Names").toDispatch();  
        Dispatch.call(Names,"Add",name,place,false).toDispatch();  
        Dispatch.put(Names, "Comment", comment); //插入备注  
    }  
    /**
     * 获取名称管理器 
     * @param name 名称管理器名 
     * @return 
     */
    public String getName(String name) {  
        Dispatch Names=Dispatch.get(workbook, "Names").toDispatch();  
        Dispatch Name=Dispatch.call(Names,"Item",name).toDispatch();  
        return Dispatch.get(Name, "Value").toString();  
    }  
    /**
     *  单元格写入值 
     * @param //sheet  被操作的sheet
     * @param position 单元格位置，如：C1 
     * @param //type 值的属性 如：value
     * @param value 
     */
    private void setValue(String position, Object value) {  
        currentSheet=this.getCurrentSheet();  
        Dispatch cell = Dispatch.invoke(currentSheet, "Range",  
                Dispatch.Get, new Object[] { position }, new int[1])  
                .toDispatch();  
        Dispatch.put(cell, "Value", value);  
        String color=this.getColor(cell);  
        this.setFont(cell,color);  
    }  
    /**
     *  设置字体 
     */
    private void setFont(Dispatch cell,String color)  
    {  
        Dispatch font=Dispatch.get(cell, "Font").toDispatch();  
        //Dispatch.put(font,"FontStyle", "Bold Italic");  
        Dispatch.put(font,"size", "1");  
        Dispatch.put(font,"color",color);  
    }  
    /**
     *  获取背景颜色 
     */
    private String getColor(Dispatch cell)  
    {  
        Dispatch Interior=Dispatch.get(cell, "Interior").toDispatch();  
        String color=Dispatch.get(Interior, "color").toString();  
        return color;  
    }  
    /**
     * 单元格读取值 
     * @param position 单元格位置，如： C1 
     * @param //sheet
     * @return 
     */
    private Variant getValue(String position) {  
        currentSheet=this.getCurrentSheet();  
        Dispatch cell = Dispatch.invoke(currentSheet, "Range", Dispatch.Get,  
                new Object[] { position }, new int[1]).toDispatch();  
        Variant value = Dispatch.get(cell, "Value");  
        return value;  
    }   
    /**
     * 获取最大行数 
     * @return 
     */
    private int getRowCount() {  
        currentSheet=this.getCurrentSheet();  
        Dispatch UsedRange=Dispatch.get(currentSheet, "UsedRange").toDispatch();  
        Dispatch rows=Dispatch.get(UsedRange, "Rows").toDispatch();  
        int num=Dispatch.get(rows, "count").getInt();  
        return num;  
    }  
    /**
     * 获取最大列数 
     * @return 
     */
    private int getColumnCount() {  
        currentSheet=this.getCurrentSheet();  
        Dispatch UsedRange=Dispatch.get(currentSheet, "UsedRange").toDispatch();  
        Dispatch Columns=Dispatch.get(UsedRange, "Columns").toDispatch();  
        int num=Dispatch.get(Columns, "count").getInt();  
        return num;  
    }  
    /**
     * 获取位置 
     * @param rnum 最大行数 
     * @param cnum 最大列数 
     */
    private String getCellPosition(int rnum,int cnum)  
    {    
          String cposition="";  
          if(cnum>26)  
          {  
              int multiple=(cnum)/26;  
              int remainder=(cnum)%26;  
              char mchar=(char)(multiple+64);  
              char rchar=(char)(remainder+64);  
              cposition=mchar+""+rchar;  
          }  
          else  
          {  
              cposition=(char)(cnum+64)+"";  
          }  
          cposition+=rnum;  
          return cposition;  
    }  
     
     /*
      * 取消兼容性检查，在保存或者另存为时改检查会导致弹窗 
      */
    public void setCheckCompatibility(){  
        Dispatch.put(workbook, "CheckCompatibility", false);  
   }  

     /* 
      *  为每个表设置打印区域 
      */  
  /*  private void setPrintArea(){  
        int count = Dispatch.get(sheets, "count").changeType(Variant.VariantInt).getInt();  
        for (int i = count; i >= 1; i--) {  
               sheet = Dispatch.invoke(sheets, "Item",  
                       Dispatch.Get, new Object[] { i }, new int[1]).toDispatch();  
           Dispatch page = Dispatch.get(sheet, "PageSetup").toDispatch();  
           Dispatch.put(page, "PrintArea", false);  
           Dispatch.put(page, "Orientation", 2);  
           Dispatch.put(page, "Zoom", false);      //值为100或false  
           Dispatch.put(page, "FitToPagesTall", false);  //所有行为一页  
           Dispatch.put(page, "FitToPagesWide", 1);      //所有列为一页(1或false)     
    }  }  */
    
}  