import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.tplink.smb.tpwp.project.export.ExportManager;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.lang.reflect.Field;
import java.util.List;

/**
 * @author Tyrian
 */

public class WordUtil {

    // debug
    private static final Logger  logger = LoggerFactory.getLogger(WordUtil.class);
    private static final boolean DEBUG  = false;

    // word document
    private Dispatch doc;

    // word运行程序对象
    private ActiveXComponent word;

    // All word document collections
    private Dispatch documents;
    private Dispatch selection;
    private boolean  saveOnExit;

    private static final String WORD_APPLICATION   = "Word.Application";
    private static final String EXCEL_APPLICATION  = "Excel.Application";
    private static final String VISIBLE            = "Visible";
    private static final String DOCUMENTS          = "Documents";
    private static final String OPEN               = "Open";
    private static final String CLOSE              = "Close";
    private static final String SAVE               = "Save";
    private static final String QUIT               = "Quit";
    private static final String SELECTION          = "Selection";
    private static final String EXECUTE            = "Execute";
    private static final String WORD_BASIC         = "WordBasic";
    private static final String FILE_SAVE_AS       = "FileSaveAs";
    private static final int    WDDONOTSAVECHANGES = 0;

    private static final String PARAM_TRUE            = "True";
    private static final String FIND                  = "Find";
    private static final String FIND_PARA_TEXT        = "Text";
    private static final String FIND_PARA_WRAP        = "Wrap";
    private static final String FIND_PARA_FORWARD     = "Forward";
    private static final String FIND_PARA_FORMAT      = "Format";
    private static final String FIND_PARA_MATCH_WHOLE = "MatchWholeWord";

    private static final String RANGE  = "Range";
    private static final String FIELDS = "Fields";
    private static final String TOC    = "TOC";
    private static final String ADD    = "Add";

    private static final String MOVE_DOWN    = "MoveDown";
    private static final String MOVE_RIGHT   = "MoveRight";
    private static final String INSERT_BREAK = "InsertBreak";
    private static final String END_KEY      = "EndKey";
    private static final String HOME_KEY     = "HomeKey";

    private static final String IN_LINE_SHAPES = "InlineShapes";
    private static final String ADD_OLE_OBJECT = "AddOLEObject";
    private static final String ADD_PICTURE    = "AddPicture";

    private static final int    BREAK_SIGN_VALUE = 7;
    private static final String WHOLE_STORY      = "WholeStory";
    private static final String UPDATE           = "Update";
    private static final String SELECT           = "Select";
    private static final String TEXT             = "Text";
    private static final String CONTENT          = "Content";
    private static final String COPY             = "Copy";
    private static final String PASTE            = "Paste";

    private static final String TYPE_PARAGRAPH   = "TypeParagraph";
    private static final String PARAGRAPH_FORMAT = "ParagraphFormat";

    // table operation
    private static final String TABLES  = "Tables";
    private static final String COLUMNS = "Columns";
    private static final String ROWS    = "Rows";
    private static final String COUNT   = "Count";
    private static final String ITEM    = "Item";
    private static final String CELL    = "Cell";
    private static final String MERGE   = "Merge";

    // style
    private static final String STYLE        = "Style";
    private static final String STYLES       = "Styles";
    private static final String TABLE_STYLE  = "table";  // table style
    private static final String HEADER_STYLE = "header";
    private static final String LIST_STYLE   = "list";
    public static final  int    LEVEL_ONE    = 1;
    public static final  int    LEVEL_TWO    = 2;
    public static final  int    LEVEL_THREE  = 3;
    public static final  int    LEVEL_FOUR   = 4;

    public static final String BODY_STYLE = "body";
    public static final String NOTE_STYLE = "note";

    private static final String WIDTH  = "Width";
    private static final String HEIGHT = "Height";

    // alignment
    private static final String ALIGNMENT    = "Alignment";
    private static final int    ALIGN_CENTER = 1;
    private static final int    ALIGN_RIGHT  = 2;
    private static final int    ALIGN_LEFT   = 3;

    //toc
    public static final String LABLE_TOC = "LABEL_TOC";

    public WordUtil() {
        this(DEBUG);
    }

    public WordUtil(boolean visible) {

        ComThread.InitSTA();
        if (word == null) {
            word = new ActiveXComponent(WORD_APPLICATION);
            word.setProperty(VISIBLE, new Variant(visible));
        }

        if (documents == null)
            documents = word.getProperty(DOCUMENTS).toDispatch();

        saveOnExit = false; // do not save when exit
    }

    public void createNewDocument() {
        doc = Dispatch.call(documents, ADD).toDispatch();
        selection = Dispatch.get(word, SELECTION).toDispatch();
    }

    public void openDocument(String docPath) {
        if (new File(docPath).exists()) {
            doc = Dispatch.call(documents, OPEN, docPath).toDispatch();
            selection = Dispatch.get(word, SELECTION).toDispatch();
        } else {
            throw new NullPointerException("doc not exists.");
        }
    }

    public void saveAs(String savePath) {
        Dispatch.call(Dispatch.call(word, WORD_BASIC).getDispatch(), FILE_SAVE_AS, savePath);
    }

    public void save() {
        Dispatch.call(this.doc, SAVE);
        Dispatch.call(this.doc, CLOSE, new Variant(true));
    }

    public void closeDocument() {
        if (doc != null) {
            Dispatch.call(doc, CLOSE, new Variant(saveOnExit));
            doc = null;
        }
    }

    public void close() {
        closeDocument();
        if (word != null) {
            Dispatch.call(word, QUIT, new Variant(WDDONOTSAVECHANGES));
            word = null;
        }
        selection = null;
        documents = null;

        ComThread.Release(); // important, otherwise may cause memory leak
    }

    /**********************
     * insert utility
     **********************/

    public void insertBodyText(String text) {
        Dispatch.put(selection, TEXT, text);

        Dispatch style = Dispatch.call(doc, STYLES, new Variant(BODY_STYLE)).toDispatch();
        Dispatch.put(selection, STYLE, style);
    }

    public void insertNote(String note) {
        Dispatch.put(selection, TEXT, note);

        Dispatch style = Dispatch.call(doc, STYLES, new Variant(NOTE_STYLE)).toDispatch();
        Dispatch.put(selection, STYLE, style);

        newLine();
    }

    /**
     * 插入标题，默认样式为 “header" + level(number).在模板中定义
     *
     * @param level 几级标题
     * @param text  标题名称
     */
    public void insertTitle(int level, String text) {
        Dispatch.put(selection, TEXT, text);

        Dispatch style = Dispatch.call(doc, STYLES, new Variant(HEADER_STYLE + level)).toDispatch();
        Dispatch.put(selection, STYLE, style);

        newLine();
    }

    /**
     * 插入无序列表
     *
     * @param text
     */
    public void insertList(String text) {
        Dispatch.put(selection, TEXT, text);
        Dispatch style = Dispatch.call(doc, STYLES, new Variant(LIST_STYLE)).toDispatch();
        Dispatch.put(selection, STYLE, style);

        newLine();
        insertBodyText("");

    }

    /**
     * 插入图片
     *
     * @param imagePath
     * @param width
     * @param height
     */
    public void insertImage(String imagePath, int width, int height) {
        Dispatch inLineShapes = Dispatch.get(selection, IN_LINE_SHAPES).toDispatch();
        Dispatch picture = Dispatch.call(inLineShapes, ADD_PICTURE, imagePath).toDispatch();

        Dispatch.call(picture, SELECT);
        Dispatch.put(picture, WIDTH, new Variant(width));
        Dispatch.put(picture, HEIGHT, new Variant(height));

        Dispatch alignment = Dispatch.get(selection, PARAGRAPH_FORMAT).toDispatch();
        Dispatch.put(alignment, ALIGNMENT, ALIGN_CENTER);

        newLine();
    }

    /**
     * 插入表格
     *
     * @param rows
     * @param cols
     * @return
     */
    public Dispatch insertTable(int rows, int cols) {
        Dispatch tables = Dispatch.get(doc, TABLES).toDispatch();
        Dispatch range = Dispatch.get(selection, RANGE).toDispatch();
        Dispatch newTable = Dispatch.call(tables, ADD, range, new Variant(rows), new Variant(cols)).toDispatch();

        Dispatch style = Dispatch.call(doc, STYLES, new Variant(TABLE_STYLE)).toDispatch();
        Dispatch.put(selection, STYLE, style);

        return newTable;

    }

    public void mergeTableCell(Dispatch table, int startRow, int startCol, int endRow, int endCol) {
        Dispatch startCell = Dispatch.call(table, CELL,
            new Variant(startRow), new Variant(startCol))
            .toDispatch();
        Dispatch endCell = Dispatch.call(table, CELL,
            new Variant(endRow), new Variant(endCol))
            .toDispatch();
        Dispatch.call(startCell, MERGE, endCell);
    }

    /**
     * 往指定的表格的单元格里填写数据
     *
     * @param table
     * @param rowIndex
     * @param colIndex
     * @param txt
     */
    public void writeTextToCell(Dispatch table, int rowIndex, int colIndex, String txt) {
        Dispatch cell = Dispatch.call(table, CELL, new Variant(rowIndex),
            new Variant(colIndex)).toDispatch();
        Dispatch.call(cell, SELECT);
        Dispatch.put(selection, TEXT, txt);
    }

    private int getColumns(Dispatch table) {
        Dispatch cols = Dispatch.get(table, COLUMNS).toDispatch();
        int colCount = Dispatch.get(cols, COUNT).changeType(Variant.VariantInt).getInt();

        return colCount;
    }

    private int getRows(Dispatch table) {
        Dispatch rows = Dispatch.get(table, ROWS).toDispatch();
        int rowCount = Dispatch.get(rows, COUNT).changeType(Variant.VariantInt).getInt();

        return rowCount;
    }

    /**
     * 表格应用样式为“table”，内置在word模板中
     *
     * @param table
     * @param datas
     * @param <T>
     */
    public <T> void insertTextToTable(Dispatch table, List<String> headers, List<T> datas) {
        if (LangUtil.listEmpty(headers))
            return;

        // write data to header
        for (int col = 0; col < headers.size(); col++) {
            writeTextToCell(table, 1, col + 1, headers.get(col));
        }

        if (LangUtil.listEmpty(datas))
            return;

        T data = datas.get(0);
        Field[] fields = data.getClass().getDeclaredFields();
        int fieldCount = fields.length;
        int columns = getColumns(table);
        if (fieldCount > columns || headers.size() > columns)
            return;

        // start from the second row
        int START_ROW = 2;
        int rowIndex;
        for (int i = 0; i < datas.size(); i++) {
            data = datas.get(i);
            rowIndex = i + START_ROW;

            for (int columnIndex = 1; columnIndex <= fieldCount; columnIndex++) {// 每一行中的单元格数
                Field field = fields[columnIndex - 1];
                try {
                    Object value = getFieldValue(data, field);
                    writeTextToCell(table, rowIndex, columnIndex, value.toString());

                } catch (Exception e) {
                    break;
                }
            }
        }
    }
	
	 public static <T> Object getFieldValue(T t, Field field) throws Exception {
        String fieldName = field.getName();
        String getMethodName =
            "get" + fieldName.substring(0, 1).toUpperCase(Locale.ROOT) + fieldName.substring(1);

        Class clazz = t.getClass();
        Method getMethod = clazz.getMethod(getMethodName, new Class[] {});
        Object value = getMethod.invoke(t, new Object[] {});

        return value;
    }

    /**
     * 插入目录
     */
    public void insertToc() {
        Dispatch range = Dispatch.get(this.selection, RANGE).toDispatch();
        Dispatch fields = Dispatch.call(this.selection, FIELDS).toDispatch();

        Dispatch.call(fields,
            ADD,
            range,
            new Variant(13),
            new Variant(TOC + "\\h"),
            new Variant(true));
    }

    public void copyContentFromOtherDoc(String docPath) {
        Dispatch otherDoc = null;
        try {
            otherDoc = Dispatch.call(documents, OPEN, docPath).toDispatch();
            Dispatch content = Dispatch.call(otherDoc, CONTENT).toDispatch();

            // copy content from part one
            Dispatch.call(content, COPY);

            Dispatch range = Dispatch.get(selection, RANGE).toDispatch();
            Dispatch.call(range, PASTE);
        } catch (Exception e) {
            logger.error("copy from other doc failed.", e);
        } finally {
            if (otherDoc != null) {
                Dispatch.call(otherDoc, CLOSE, new Variant(false));
            }
        }
    }

    /**
     * 从选定内容或插入点开始查找文本
     *
     * @param findContent 要查找的文本
     * @return boolean true-查找到并选中该文本，false-未查找到文本
     */
    public boolean find(String findContent) {
        if (StringUtil.isEmpty(findContent))
            return false;

        Dispatch find = word.call(selection, FIND).toDispatch();
        Dispatch.put(find, FIND_PARA_TEXT, findContent);
        Dispatch.put(find, FIND_PARA_FORWARD, PARAM_TRUE);
        Dispatch.put(find, FIND_PARA_FORMAT, PARAM_TRUE);
        Dispatch.put(find, FIND_PARA_WRAP, Integer.valueOf(1));
        Dispatch.put(find, FIND_PARA_MATCH_WHOLE, PARAM_TRUE);

        return Dispatch.call(find, EXECUTE).getBoolean();
    }

    public boolean findAndReplace(String findContent, String replaceContent) {
        boolean found = find(findContent);

        if (found) {
            Dispatch.put(this.selection, FIND_PARA_TEXT, replaceContent);
        }
        return found;
    }

    public void replaceAllText(String toFindText, String newText) {
        while (find(toFindText)) {
            Dispatch.put(selection, TEXT, newText);
            Dispatch.call(selection, MOVE_RIGHT);
        }
    }

    public Dispatch getTable(Dispatch doc, int index) {
        Dispatch tables = Dispatch.get(doc, TABLES).toDispatch();
        Dispatch table = Dispatch.call(tables, "Item", new Variant(index))
            .toDispatch();

        return table;
    }

    public void pasteFileAsObject(String filePath, String objectName, String iconPath) {
        if (!FileUtil.isFileExist(filePath)) {
            return;
        }

        ActiveXComponent xlsApp = null;
        Dispatch workbook = null;
        try {
            xlsApp = new ActiveXComponent(EXCEL_APPLICATION);

            xlsApp.setProperty("CutCopyMode", new Variant(true));
            xlsApp.setProperty(VISIBLE, new Variant(false));
            Dispatch workbooks = xlsApp.getProperty("Workbooks").toDispatch();
            workbook = Dispatch.invoke(workbooks, OPEN, 1,
                new Object[] { filePath, new Variant(false), new Variant(true) }, new int[1])
                .toDispatch();

            Dispatch sheets = Dispatch.get(workbook, "Sheets").toDispatch();

            Dispatch.call(sheets, SELECT, new Variant(true));
            Dispatch selectContent = Dispatch.get(xlsApp, SELECTION).toDispatch();
            Dispatch.call(selectContent, COPY);

            Dispatch alignment = Dispatch.get(this.selection, PARAGRAPH_FORMAT).toDispatch();
            Dispatch.put(alignment, ALIGNMENT, ALIGN_CENTER);
            Dispatch.call(this.selection, "PasteSpecial", new Variant(1), new Variant(false), new Variant(0),
                new Variant(true), new Variant(0), new Variant(iconPath), new Variant(objectName));

            selectContent.safeRelease();
            sheets.safeRelease();
            workbooks.safeRelease();
            xlsApp.setProperty("CutCopyMode", new Variant(false));
        } finally {
            if (workbook != null) {
                Dispatch.call(workbook, CLOSE, new Variant(false));
                workbook.safeRelease();
            }

            if (xlsApp != null) {
                Dispatch.call(xlsApp, QUIT);
                xlsApp.safeRelease();
            }
        }
    }

    public void insertExcelAsObject(String filePath, String objectName) {
        if (!FileUtil.isFileExist(filePath))
            return;

        String excelIcon = FileUtil.getFile("template" + File.separator + "excel.ico").getAbsolutePath();
        pasteFileAsObject(filePath, objectName, excelIcon);
    }

    public void insertFileAsObject(String filePath, String objectName, String insertType, String iconPath) {
        if (!FileUtil.isFileExist(filePath)) {
            return;
        }

        Dispatch shapes = Dispatch.get(this.selection, IN_LINE_SHAPES).toDispatch();
        Dispatch alignment = Dispatch.get(this.selection, PARAGRAPH_FORMAT).toDispatch();
        Dispatch.put(alignment, ALIGNMENT, ALIGN_CENTER);
        Dispatch.call(shapes, ADD_OLE_OBJECT, new Variant(insertType), new Variant(filePath),
            new Variant(false), new Variant(true), new Variant(iconPath), new Variant(1),
            new Variant(objectName)).toDispatch();
    }

    /*********************
     * navigation utility
     *********************/
    public void moveStart() {
        Dispatch.call(selection, HOME_KEY, new Variant(6));
    }

    public void backspace(int times) {
        if (times < 1)
            throw new IllegalArgumentException("Invalid backspace times");

        for (int i = 0; i < times; i++) {
            Dispatch.call(selection, "TypeBackspace");
        }
    }

    public void newPage() {
        moveEnd();
        insertBreak();
    }

    public void newLine() {
        moveRight();
        Dispatch.call(selection, TYPE_PARAGRAPH);
    }

    public void moveEnd() {
        Dispatch.call(selection, END_KEY, new Variant(6));
    }

    public void insertBreak() {
        Dispatch.call(selection, INSERT_BREAK, new Variant(7));
    }

    public void moveDown() {
        Dispatch.call(selection, MOVE_DOWN);
    }

    private void moveRight() {
        Dispatch.call(selection, MOVE_RIGHT, new Variant(1), new Variant(1));
    }

    public static void main(String[] args) {

        WordUtil template = new WordUtil(true);
        template.createNewDocument();

        template.newPage();
        template.insertList("cccccccccc");
        template.insertList("cccccccccc");
        template.insertList("cccccccccc");

        template.insertNote("this is note");
        template.newLine();
        template.insertTitle(LEVEL_FOUR, "level four title");
        template.insertBodyText("boddyddsddfsfsf");
        template.newLine();
        template.newLine();

        template.insertBodyText("excel");
        template.findAndReplace("excel", "");

        Dispatch table = template.insertTable(3, 8);
        template.mergeTableCell(table, 1, 1, 1, 8);
        template.writeTextToCell(table, 1, 1, "dfsfsfsfsfs");
        template.writeTextToCell(table, 2, 1, "111");

        template.saveAs("test.doc");
        template.close();
    }
}
