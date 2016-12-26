package jp.S.common.tcp;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.PathMatcher;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Optional;
import java.util.Set;
import java.util.TreeMap;
import java.util.TreeSet;
import java.util.concurrent.ConcurrentHashMap;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * TcpClientUtilを使用する業務APのUT/ITa用スタブ。
 *
 * ■stubResponseMap: Map<[電文ID:String], [電文ID毎にテストデータをため込んだマップ:Map]>
 *   └電文ID毎～マップ: Map<[電文ID毎のキー項目値(memId):String], レスポンス用マップ:Map>
 *     └応答電文用マップ: Map<[電文項目名:String], [電文項目値:Object]>
 *
 * ■stubKeyMap: Map<[電文ID:String], [電文ID毎のキー項目名リスト:List<電文ID毎のキー項目名>]>
 *
 * ■argumentMap: Map<[電文ID:String], [電文ID毎に呼出順で引数のリクエストマップを溜め込むリスト:List<>]>
 *   └電文ID毎に呼出順で引数のリクエストマップを溜め込むリスト:List<[リクエストマップ:Map]>
 *
 * @since 1.0
 */
public class TcpClientStub implements TcpClient {

    /** logger */
    private static final Logger LOGGER = LoggerManager.get(TcpClientStub.class);

    /** LIST_MAPの判定用 */
    private static final String LIST_MAP_KEY = "LIST_MAP";

    /** keyMap生成用固定キー値 */
    private static final String SOCKET_RESPONSE_KEY = "socketResponseKey";

    /** UT用スタブデータファイル拡張子 */
    private static final String BOOK_EXTENSIONS = ".xlsx:.xls";

    /** 取引単体・IT用スタブデータファイル拡張子 */
    private static final String IT_BOOK_EXTENSIONS = "regex:.+\\.xlsx?";

    /** スタブ設定 */
    private static TcpClientStubSettings settings = new TcpClientStubSettings();

    /** デフォルトデータ用値設定Map<電文ID, 応答電文Map特定キーMap<応答電文Map特定キー, 応答電文Map<項目名, 項目値>>> */
    private static Map<String, Map<String, Map<String, Object>>> defaultResponseMap = new ConcurrentHashMap<>();

    /** デフォルトデータ最終更新日時 */
    private static long defaultResponseBookLastModified = -1L;

    /** デフォルトデータ用ロックオブジェクト */
    private static Object defaultResponseLock = new Object();

    /** スタブデータ用値設定Map<電文ID, 応答電文Map特定キーMap<応答電文Map特定キー, 応答電文Map<項目名, 項目値>>> */
    private static Map<String, Map<String, Map<String, Object>>> stubResponseMap = new ConcurrentHashMap<>();

    /** スタブデータ用キー項目Map<電文ID, 全ブック分のキー項目名リスト<キー項目名List<キー項目名>>> */
    private static Map<String, Set<List<String>>> stubKeyMap = new ConcurrentHashMap<>();

    /** スタブデータ最終更新日時（ファイル毎に保持） */
    private static Map<File, Long> stubResponseBooksLastModified = new ConcurrentHashMap<>();

    /** スタブデータ用ロックオブジェクト */
    private static Object stubResponseLock = new Object();

    /** UTフラグ(呼出時の引数Mapをため込むか否かの判定に使用する) */
    private static boolean isUT = false;

    /** UT用の前回テストクラスインスタンス */
    private static Object lastTestClass = new Object();

    /** UT用の呼出時の引数Map */
    private static Map<String, List<Map<String, Object>>> argumentMap = new ConcurrentHashMap<>();

    /** フォーマッタファクトリ */
    private static FormatterFactory formatterFactory = new FormatterFactory();

    /**
     * UT用の初期化処理（Excelからテストデータを読み込んで溜め込む）
     * 制約として、Excelファイルの配置場所はテストFWの標準パスのみとする。
     * （N.test.resource-rootによるベースパスは利用できない）
     *
     * @param testClass テストクラスオブジェクト
     * @param utSheetName UTで使用するシート名
     */
    public static void initialize(Object testClass, String utSheetName) {
        if (!lastTestClass.getClass().getName().equals(testClass.getClass().getName())) {
            stubResponseBooksLastModified = new ConcurrentHashMap<>();
            setUTBookPathName(testClass);
        }
        if (!settings.stubResponseSheetName.equals(utSheetName)) {
            stubResponseBooksLastModified = new ConcurrentHashMap<>();
            settings.stubResponseSheetName = utSheetName;
        }

        initialize();
        clearArgumentMap();

        isUT = true;
    }

    /**
     * 初期化処理（Excelからテストデータを読み込んで溜め込む）
     * （取引単体・ITa用）
     */
    public static void initialize() {

        // デフォルトのレスポンス情報を初期化
        if (isDefaultBookUpdated()) {
            synchronized(defaultResponseLock ) {
                initializeDefaultResponse();
            }
        }

        // スタブのレスポンス情報を初期化
        if (isAnyStubBookUpdated()) {
            synchronized(stubResponseLock) {
                initializeStubResponse();
            }
        }
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public Map<String, Object> sendSync(String telegramId, Map<String, Object> requestMap) {

        // UT用に引数をキャッシュ
        cashArgumentMap(telegramId, requestMap);

        // テスト仕様書の更新日時が変わっていたら再初期化する
        initialize();

        // 電文IDと要求電文Mapをキーに、応答電文Mapを取得して返す
        return getResponseMap(telegramId, requestMap);
    }

    /**
     * UTの場合は、要求電文Mapを整形してキャッシュする。
     * @param telegramId 電文ID
     * @param requestMap 要求電文Map
     */
    private void cashArgumentMap(String telegramId, Map<String, Object> requestMap) {

        if (isUT) {
            if (!argumentMap.containsKey(telegramId)) {
                argumentMap.put(telegramId, new ArrayList<Map<String, Object>>());
            }
            argumentMap.get(telegramId).add(shapeRequestMap(requestMap));
        }
    }

    /**
     * 要求電文Mapを呼出時引数Map用に整形する。
     *   ・byte配列項目のassertが可能となるようHex表記に変換する。
     *     ⇒BINARY(0x00000000)
     *
     * @param requestMap 要求電文Map
     * @return 要求電文Mapを整形したMap
     */
    private Map<String, Object> shapeRequestMap(Map<String, Object> requestMap) {
        Map<String, Object> shapedMap = new TreeMap<>();

        for (Entry<String, Object> entry : requestMap.entrySet()) {
            if (entry.getValue() instanceof byte[]) {
                shapedMap.put(entry.getKey(), String.format("BINARY(%s)",
                        BinaryUtil.convertToHexStringWithPrefix((byte[]) entry.getValue())));
            } else {
                shapedMap.put(entry.getKey(), entry.getValue());
            }
        }
        return shapedMap;
    }

    /**
     * (UT用)コンポーネント呼出時の引数を取得する。
     * コンポーネント呼出時の引数Mapを、呼び出した順番でリスト化したオブジェクトを取得する。
     * @param telegramId 電文ID
     * @return 引数Mapリスト
     */
    public static List<Map<String, Object>> getArgumentMapList(String telegramId) {
        return argumentMap.get(telegramId);
    }

    /**
     * 設定されたスタブ用マップを初期化する。
     */
    public static void clear() {

        stubResponseMap.clear();
        stubKeyMap.clear();
        stubResponseBooksLastModified = new ConcurrentHashMap<>();

        defaultResponseMap.clear();
        defaultResponseBookLastModified = -1L;

        clearArgumentMap();
        isUT = false;
    }

    /**
     * (UT用)コンポーネント呼出時の引数Mapリストを初期化する。
     */
    public static void clearArgumentMap() {
        argumentMap.clear();
    }

    /**
     * テスト対象のデフォルトデータが更新されたか否かを取得する。
     * @return テスト対象のデフォルトデータが更新された場合:true 左記以外の場合:false
     */
    private static boolean isDefaultBookUpdated() {
        return defaultResponseBookLastModified != new File(
                settings.defaultResponseBookPath + settings.defaultResponseBookName).lastModified();
    }

    /**
     * スタブのレスポンス情報を初期化する。
     */
    private static void initializeDefaultResponse() {
        defaultResponseMap = Collections.emptyMap();

        // ワークブックを取得
        Workbook defaultResponseBook = openDefaultResponseBook();
        if (defaultResponseBook == null) {
            LOGGER.logWarn(String.format(
                    "default socket response workbook doesn't exist.[%s]",
                    settings.defaultResponseBookPath + settings.defaultResponseBookName));
            return;
        }

        // ワークシートを取得
        Sheet defaultResponseSheet = defaultResponseBook.getSheet(settings.defaultResponseSheetName);
        if (defaultResponseSheet == null) {
            LOGGER.logWarn(String.format(
                    "default socket response worksheet doesn't exist.[%s]", settings.defaultResponseSheetName));
            return;
        }

        // テスト仕様書, sheetName からkeyMapにキーを溜め込む
        Map<String, List<String>> defaultKeyMap = createKeyMap(defaultResponseSheet);
        if (defaultKeyMap == null) {
            LOGGER.logWarn("default socket response worksheet doesn't have [LIST_MAP=socketResponseKey].");
            return;
        }

        // テスト仕様書, sheetName, keyMap からdefaultResponseMapに戻り値のマップを溜め込む
        defaultResponseMap = createResponseMap(defaultResponseSheet, defaultKeyMap);
    }

    /**
     * デフォルトデータ用ブックを開く。
     * デフォルトデータ用ブックを開いた時点の更新日時を保持する。
     *
     * @return デフォルトデータ用ワークブック
     */
    private static Workbook openDefaultResponseBook() {
        File file = new File(settings.defaultResponseBookPath + settings.defaultResponseBookName);
        if (!file.exists()) {
            return null;
        }
        defaultResponseBookLastModified = file.lastModified();

        return openWorkBook(file);
    }

    /**
     * Excelワークブックを開く。
     * @param file ワークブックファイル
     * @return ワークブック
     */
    private static Workbook openWorkBook(File file) {
        String absolutePath = file.getAbsolutePath();
        LOGGER.logDebug("opening file:" + absolutePath);

        try (InputStream in = new FileInputStream(absolutePath)) {
            return WorkbookFactory.create(in);
        } catch (IOException | InvalidFormatException e) {
            throw new IllegalStateException(String.format(
                    "file open error.[%s]", file.getAbsolutePath()), e);
        }
    }

    /**
     * テスト対象のスタブデータファイルがどれか一つでも更新されたか否かを取得する。
     * @return 一つでも更新された場合:true 左記以外の場合:false
     */
    private static boolean isAnyStubBookUpdated() {
        return getStubResponsePathStream().count() != stubResponseBooksLastModified.size() ||
                getStubResponsePathStream().anyMatch((TcpClientStub::isStubBookUpdated));
    }

    /**
     * 対象のスタブデータが更新されたか否かを取得する。
     * @param path スタブデータファイルのパス
     * @return テスト対象のスタブデータが更新された場合:true 左記以外の場合:false
     */
    private static boolean isStubBookUpdated(Path path) {
        File file = path.toFile();
        Long lastModified = stubResponseBooksLastModified.get(file);
        return lastModified == null || lastModified.longValue() != file.lastModified();
    }

    /**
     * スタブのレスポンス情報を初期化する。
     */
    private static void initializeStubResponse() {
        stubKeyMap = new ConcurrentHashMap<>();
        stubResponseMap = new ConcurrentHashMap<>();
        stubResponseBooksLastModified = new ConcurrentHashMap<>();

        // 各ワークブックを初期化
        getStubResponsePathStream().forEach(TcpClientStub::initializeStubResponseUnit);
    }

    /**
     * スタブデータファイルのストリームを返却する。
     * @return スタブデータのファイルパス（ストリーム）
     */
    private static Stream<Path> getStubResponsePathStream() {
        Path path = Paths.get(settings.stubResponseBookPath + settings.stubResponseBookName);
        PathMatcher filter = path.getFileSystem().getPathMatcher(IT_BOOK_EXTENSIONS);
        if (path.toFile().isDirectory()) {
            return getFilePathStream(path).filter(filter::matches);
        }
        return  Arrays.stream(new Path[]{path});
    }

    /**
     * 指定したディレクトリパス直下の全ファイル(ディレクトリ,隠しファイル除く)のストリームを返却する。
     * ※Excelが作成する~$socketResponse.xlsを抑止することを目的とする。
     * @param path ディレクトリパス
     * @return パスストリーム
     */
    private static Stream<Path> getFilePathStream(Path path) {
        try (Stream<Path> pathStream = Files.list(path)) {
            return pathStream.filter(p -> !p.toFile().isDirectory() && !p.toFile().isHidden()).collect(
                    Collectors.toList()).stream(); // DirectoryStreamをcloseするため、順次Streamに置き換えて返却
        } catch (IOException ioe) {
            throw new IllegalStateException(String.format("path access error.[%s]", path.toString()), ioe);
        }
    }

    /**
     * スタブのレスポンス情報を初期化する（ファイル単体）。
     * @param stubResponseBookFile スタブデータファイルのパス
     * @return 正常にスタブデータを取得できた場合true、それ以外の場合はfalse
     */
    private static void initializeStubResponseUnit(Path stubResponseBookFile) {
        if (!stubResponseBookFile.toFile().exists()) {
            LOGGER.logWarn(String.format(
                    "unit test workbook doesn't exist.[%s]", stubResponseBookFile.toString()));
            return;
        }

        Workbook stubResponseBook = openStubResponseBook(stubResponseBookFile);

        // ワークシートを取得
        Sheet stubResponseSheet = stubResponseBook.getSheet(settings.stubResponseSheetName);
        if (stubResponseSheet == null) {
            LOGGER.logWarn(String.format(
                    "unit test worksheet doesn't exist.[%s]", settings.stubResponseSheetName));
            return;
        }

        // テスト仕様書, sheetName からkeyMapにキーを溜め込む
        Map<String, List<String>> tmpKeyMap = createKeyMap(stubResponseSheet);
        if (tmpKeyMap == null) {
            return;
        }
        tmpKeyMap.entrySet().stream().forEach(TcpClientStub::updateStubKeyMap);

        // テスト仕様書, sheetName, keyMap からstubResponseMapに戻り値のマップを溜め込む
        createResponseMap(stubResponseSheet, tmpKeyMap).entrySet().forEach(TcpClientStub::reflectToStubResponseMap);
    }

    /**
     * スタブデータに指定されたエントリーを反映する。
     * @param entry エントリー<電文ID, <応答電文Map特定キー, 応答電文Map<項目名, 項目値>>>
     */
    private static void reflectToStubResponseMap(Map.Entry<String, Map<String, Map<String, Object>>> responseMapEntry) {
        if (!stubResponseMap.containsKey(responseMapEntry.getKey())) {
            stubResponseMap.put(responseMapEntry.getKey(), new HashMap<>());
        }
        Map<String, Map<String, Object>> specifyingKeyMap = stubResponseMap.get(responseMapEntry.getKey());
        specifyingKeyMap.putAll(responseMapEntry.getValue());
    }

    /**
     * スタブデータ用ブックを開く。
     * 開いた時点の更新日時を保持する。
     * @param path  スタブデータ用ブックのパス
     * @return スタブデータ用ワークブック
     */
    private static Workbook openStubResponseBook(Path path) {
        File file = path.toFile();
        stubResponseBooksLastModified.put(file, Long.valueOf(file.lastModified()));
        return openWorkBook(file);
    }

    /**
     * スタブデータ用キー項目MapにsocketResponseKey(電文IDごとの応答電文Map特定キー)を追加する。
     * ブック間で同一電文が定義され、かつ、socketResponseKeyが異なることを想定して、
     * 電文IDごとに複数のsocketResponseKeyを持たせる。
     * @param entry エントリー(電文ID, 電文IDに紐付くsocketResponseKey)
     */
    private static void updateStubKeyMap(Map.Entry<String, List<String>> entry) {
        Set<List<String>> telegramIdList = stubKeyMap.get(entry.getKey());
        if (telegramIdList == null) {
            Comparator<List<String>> comparator = Comparator.comparing(List::size);
            telegramIdList = new TreeSet<>(comparator.reversed()); // キーが多い(マッチ条件が厳しい)順にソート
            stubKeyMap.put(entry.getKey(), telegramIdList);
        }
        telegramIdList.add(entry.getValue());
    }

    /**
     * UTのテスト仕様書パスを設定する。
     * @param testClass テストクラス
     */
    private static void setUTBookPathName(Object testClass) {
        File file = null;
        final String innerClassMark = "$";
        final String[] extensions = BOOK_EXTENSIONS.split(":");
        String pathBase = testClass.getClass().getName();

        // インナークラス名になっている場合は本来のクラス名を取得し直す
        if (pathBase.contains(innerClassMark)) {
            pathBase = pathBase.substring(0, pathBase.indexOf(innerClassMark));
        }
        pathBase = settings.utDefaultResourcesRoot + pathBase.replace('.', '/');

        for (String ext : extensions) {
            file = new File(pathBase + ext);

            if (file.exists()) {
                settings.stubResponseBookPath = file.getParent() + File.separatorChar;
                settings.stubResponseBookName = file.getName();
                return;
            }
        }
        throw new IllegalStateException(String.format("can't get TestData Excel File. [%s]",
                file.getParent() + File.separatorChar + file.getName()));
    }

    /**
     * key情報を生成。
     * @param sheet ワークシート
     * @return socketResponseKey情報（シートに存在しない場合はnull）
     */
    private static Map<String, List<String>> createKeyMap(Sheet sheet) {

        // SOCKET_RESPONSE_KEY のセルを取得。存在しない場合はデータの記載されたシートではないと見なす
        Cell resKeyCell = getResKeyCell(sheet);
        if (resKeyCell == null) {
            return null;
        }

        // SOCKET_RESPONSE_KEYの直下から電文IDごとのキー項目を検索
        Map<String, List<String>> resultKeyMap = new ConcurrentHashMap<>();
        for (int i = resKeyCell.getRowIndex() + 1; i < sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);

            // 空行を発見したらSOCKET_RESPONSE_KEY情報の取得完了と見なす
            if (isEmptyRow(row)) {
                break;
            }

            // コメント（"//"）行は読みとばす
            if (isCommentRow(row)) {
                continue;
            }

            // KEY情報読み取り
            // 1列目:電文ID
            String telegramId = getStringCellValue(row.getCell(resKeyCell.getColumnIndex())).trim();
            List<String> keyList = new ArrayList<>();

            // 2列目以降:空セルorコメントセルを見付けるまで、キー項目名として取得
        createKeylistLoop:
            for (int j = resKeyCell.getColumnIndex() + 1; j < row.getLastCellNum(); j++) {
                String keyName = getStringCellValue(row.getCell(j)).trim();
                if (keyName.isEmpty() || keyName.startsWith("//")) {
                    break createKeylistLoop;
                }
                keyList.add(keyName);
            }
            resultKeyMap.put(telegramId, keyList);
        }
        return resultKeyMap;
    }

    /**
     * 空行判定。
     * @param row Excel行
     * @return 空行の場合:true、空行以外の場合:false
     */
    private static boolean isEmptyRow(Row row) {
        if (row != null) {
            for (int j = 0; j < row.getLastCellNum(); j++) {
                String cellValue = getStringCellValue(row.getCell(j)).trim();

                if (!cellValue.isEmpty()) {
                    return false;
                }
            }
        }
        return true;
    }

    /**
     * コメント行判定。
     * @param row Excel行
     * @return コメント行の場合:true、コメント行以外の場合:false
     */
    private static boolean isCommentRow(Row row) {
        for (int j = 0; j < row.getLastCellNum(); j++) {
            String cellValue = getStringCellValue(row.getCell(j)).trim();

            if (!cellValue.isEmpty()) {
                if (cellValue.startsWith("//")) {
                    return true;
                }
                return false;
            }
        }
        return false;
    }

    /**
     * @param sheet シート
     * @return SOCKET_RESPONSE_KEYセル
     */
    private static Cell getResKeyCell(Sheet sheet) {

        // key情報の位置特定
        final String keyMapId = LIST_MAP_KEY + "=" + SOCKET_RESPONSE_KEY;

        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);

            if (row != null) {
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    Cell cell = row.getCell(j);
                    if (keyMapId.equals(getStringCellValue(cell).replaceAll(" ", ""))) {
                        return cell;
                    }
                }
            }
        }
        return null;
    }

    /**
     * レスポンス情報を生成。
     * @param sheet ワークシート
     * @param keyMap キー情報Map
     * @return SOCKET_RESPONSE_KEY情報
     */
    private static Map<String, Map<String, Map<String, Object>>> createResponseMap(
            Sheet sheet, Map<String, List<String>> keyMap) {
        Map<String, Map<String, Map<String, Object>>> resultMap = new ConcurrentHashMap<>();

        int lastRow = sheet.getLastRowNum();

        // 電文IDごとレスポンス表を検索
        // for (int i = 0; i < lastRow; i++) {
        int i = -1; // Modified Control Variable回避(for文⇒while文)
        while (i++ < lastRow) {

            // [LIST_MAP=電文ID]が存在するまで探索
            Cell telegramIdCell = getTelegramIdCell(sheet.getRow(i), keyMap);
            if (telegramIdCell == null) {
                continue;
            }

            // 電文ID単位の応答電文Mapを生成
            String telegramId = getStringCellValue(telegramIdCell).split("=")[1];
            Map<String, Map<String, Object>> telegramIdMap = new HashMap<>();
            resultMap.put(telegramId, telegramIdMap);

            // [LIST_MAP=電文ID]の直下から電文IDごとのスタブ用値設定Mapを取得
        findItemNameLoop:
            while (i++ < lastRow) {

                // 項目物理名定義行が存在するまで探索
                Row itemNamesRow = sheet.getRow(i);
                // 空行・コメント（"//"）行は読みとばす
                if (isEmptyRow(itemNamesRow) || isCommentRow(itemNamesRow)) {
                    continue findItemNameLoop;
                }

                // 項目物理名定義行から項目名配列を取得
                String[] itemNames = getItemNames(itemNamesRow);

                // 項目物理名定義行の直下から電文IDごとのスタブ用値設定Mapを取得
            createStubMapLoop:
                while (i++ < lastRow) {
                    Row row = sheet.getRow(i);

                    // 空行か次の[LIST_MAP=電文ID]を発見したら当該電文IDのスタブ用値設定Mapの取得完了と見なす
                    if (isEmptyRow(row)) {
                        break findItemNameLoop;
                    }
                    if (getTelegramIdCell(row, keyMap) != null) {
                        i--; // 次の[LIST_MAP=電文ID]を発見した場合は、シーク行を戻してから次のレスポンス表取得処理に移行する
                        break findItemNameLoop;
                    }

                    // コメント（"//"）行は読みとばす
                    if (isCommentRow(row)) {
                        continue createStubMapLoop;
                    }

                    // 応答電文Mapを生成
                    Map<String, Object> responseMap = new HashMap<>(itemNames.length);
                    for (int j = 0; j < itemNames.length; j++) {
                        responseMap.put(itemNames[j], getCellValue(row.getCell(j)));
                    }
                    telegramIdMap.put(createKey(keyMap.get(telegramId), responseMap), responseMap);
                }
            }
        }
        return resultMap;
    }

    /**
     * 項目名を配列で取得する。
     * @param row 取得対象行
     * @return 項目名配列
     */
    private static String[] getItemNames(Row row) {
        List<String> list = new ArrayList<>();
        for (int j = 0; j < row.getLastCellNum(); j++) {
            String value = getStringCellValue(row.getCell(j)).trim();

            // 空セルまたはコメントセルに遭遇したら処理を抜ける
            if (value.isEmpty() || value.startsWith("//")) {
                break;
            }
            list.add(value);
        }
        return list.toArray(new String[list.size()]);
    }

    /**
     * セル値を取得する。
     * @param cell 値を取得する対象セル
     * @return セルに入力された内容を元に生成した値
     */
    private static Object getCellValue(Cell cell) {
        String value = getStringCellValue(cell);
        String trimedValue = value.trim();

        // SPACE(長さ)
        if (trimedValue.matches("SPACE\\(\\d+\\)")) {
            return space(Integer.parseInt(trimedValue.substring(6, trimedValue.length() - 1)));
        }

        // NUMBER(数値)
        if (trimedValue.matches("NUMBER\\(\\d+\\)")) {
            return new BigDecimal(trimedValue.substring(7, trimedValue.length() - 1));
        }

        // BINARY(バイナリ表現)
        if (trimedValue.matches("BINARY\\(0[xX]([0-9a-fA-F]{2})+\\)")) {

            return BinaryUtil.convertToBytes(trimedValue.substring(7, trimedValue.length() - 1), null);
        }
        return value;
    }

    /**
     * 指定された桁数分の半角スペース文字列を返却する。
     * @param keta 桁数
     * @return 半角スペース文字列
     */
    private static String space(int keta) {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < keta; i++) {
            sb.append(' ');
        }
        return sb.toString();
    }

    /**
     * 指定した行に[LIST_MAP=電文ID]のセルが存在する場合は取得する。
     *
     * @param row 探索対象行
     * @param keyMap キー情報Map
     * @return [LIST_MAP=電文ID]が記載されたセル
     */
    private static Cell getTelegramIdCell(Row row, Map<String, List<String>> keyMap) {
        if (row != null) {
            for (int j = 0; j < row.getLastCellNum(); j++) {
                Cell cell = row.getCell(j);
                if (cell == null) {
                    continue;
                }

                String cellValue = getStringCellValue(cell).replaceAll(" ", "");
                if (cellValue.startsWith("//")) { // コメント行は飛ばす
                    return null;
                }

                String[] cellValueArray = getStringCellValue(cell).trim().split("=");
                if (cellValueArray.length == 2 && cellValueArray[0].equalsIgnoreCase(LIST_MAP_KEY)) {

                    for (String telegramId : keyMap.keySet()) {
                        if (cellValueArray[1].equalsIgnoreCase(telegramId)) {
                            return cell;
                        }
                    }
                }
                return null;
            }
        }
        return null;
    }

    /**
     * コンポーネントの戻り値（応答電文Map）を取得する。
     *
     * @param telegramId 戻り値に紐付く電文ID
     * @param requestMap 要求電文Map
     * @return 電文IDに紐付く戻り値
     */
    private static Map<String, Object> getResponseMap(String telegramId, Map<String, Object> requestMap) {
        LOGGER.logInfo("要求データ[" + telegramId + "S]：[" + requestMap.toString() + "]");

        IllegalStateException exception = null;
        // スタブデータを取得
        Map<String, Object> responseMap = createStubResponseMap(telegramId, requestMap);

        // 紐付くスタブデータが存在しない場合は、デフォルトデータを取得
        if (responseMap == null) {
            responseMap = createDefaultResponseMap(telegramId, requestMap);
            exception = checkExceptionCase(responseMap);

        } else {
            exception = checkExceptionCase(responseMap);

            // スタブデータが存在する場合(=当該電文IDがテスト対象の場合)のみ、フォーマット定義ファイルをチェック
            checkFormatterFile(telegramId, requestMap, responseMap);
        }

        if (exception != null) {
            LOGGER.logInfo("Socket通信の応答として例外送出データが指定されました。", exception);
            throw exception;
        }
        LOGGER.logInfo("応答データ[" + telegramId + "R]：[" + responseMap.toString() + "]");
        return responseMap;
    }

    /**
     * スタブデータから応答電文Mapを取得する。
     * @param telegramId 電文ID
     * @param requestMap 要求電文Map
     * @return スタブデータから取得した応答電文Map
     */
    private static Map<String, Object> createStubResponseMap(String telegramId, Map<String, Object> requestMap) {
        if (!stubResponseMap.containsKey(telegramId)) {
            return null;
        }
        Map<String, Object> responseDataMap = searchByStubResponseMap(telegramId, requestMap);
        if (responseDataMap == null) {
            return null;
        }

        // スタブデータに存在しない項目は要求電文Mapの値を設定
        Map<String, Object> resultMap = new HashMap<>();
        resultMap.putAll(requestMap);
        for (Map.Entry<String, Object> entry : responseDataMap.entrySet()) {
            if (hasValue(entry.getValue())) {
                resultMap.put(entry.getKey(), entry.getValue());
            }
        }
        return resultMap;
    }

    /**
     * 要求電文Mapをもとにスタブデータから該当する応答電文Mapを探して返却する。
     * ※キーが重複して登録されている場合は、最初にマッチしたデータが優先される。
     * @param telegramId 電文ID
     * @param requestMap 要求電文Map
     * @return 応答電文Map
     */
    private static Map<String, Object> searchByStubResponseMap(String telegramId, Map<String, Object> requestMap) {
        Map<String, Map<String, Object>> responseDataMap =stubResponseMap.get(telegramId);
        Set<List<String>> listByBook = stubKeyMap.get(telegramId);
        List<String> candidateList = new ArrayList<>();

        listByBook.forEach(keyList -> candidateList.add(createKey(keyList, requestMap)));
        Optional<String> key = candidateList.stream().filter(responseDataMap::containsKey).findFirst();

        if (key.isPresent()) {
            return responseDataMap.get(key.get());
        }
        return null;
    }

    /**
     * フォーマット定義ファイルをチェックする。
     *   - ファイル存在チェック（要求電文/応答電文）
     *   - フォーマット定義構文チェック（要求電文/応答電文）
     *   - 要求電文出力チェック（要求電文のみ）
     *
     * @param telegramId 電文ID
     * @param requestMap 要求電文Map
     * @param responseMap 応答電文Map
     */
    private static void checkFormatterFile(String telegramId,
            Map<String, Object> requestMap, Map<String, Object> responseMap) {
        final String sendSuffix = "S"; // 要求電文の電文ID接尾辞
        final String recvSuffix = "R"; // 応答電文の電文ID接尾辞

        final File formatterFileS = FilePathSetting.getInstance().getFileWithoutCreate(
                "format", telegramId + sendSuffix);
        if (!formatterFileS.exists()) {
            throw new IllegalArgumentException("要求電文のフォーマット定義ファイルが存在しません。 "
                    + String.format("要求電文フォーマット定義ファイル:[%s]", formatterFileS.getAbsolutePath()));
        }

        final File formatterFileR = FilePathSetting.getInstance().getFileWithoutCreate(
                "format", telegramId + recvSuffix);
        if (!formatterFileR.exists()) {
            throw new IllegalArgumentException("応答電文のフォーマット定義ファイルが存在しません。 "
                    + String.format("応答電文フォーマット定義ファイル:[%s]", formatterFileR.getAbsolutePath()));
        }

        // フォーマット定義ファイルを取得し、アップロードファイルを読み込むためのフォーマッターを生成する。
        try (ByteArrayOutputStream out = new ByteArrayOutputStream();
                DataRecordFormatter formatter = formatterFactory.createFormatter(
                        formatterFileS).setOutputStream(out).initialize()) {
            formatter.writeRecord(requestMap);
            LOGGER.logInfo("要求電文[" + telegramId + sendSuffix + "]：[" + Base64Util.encode(out.toByteArray()) + "]");
        } catch (InvalidDataFormatException | IOException e) {
            throw new InvalidDataFormatException("要求電文マップで項目が不足しているか、"
                    + "要求電文マップの項目値がフォーマット定義ファイルで定義される項目の型に違反しています。", e);
        } catch (Exception e) {
            throw new InvalidDataFormatException("要求電文フォーマット定義ファイルの構文に誤りがあります。", e);
        }

        try (DataRecordFormatter formatter = formatterFactory.createFormatter(formatterFileR);
                ByteArrayInputStream in = new ByteArrayInputStream(new byte[0])) {
            formatter.setInputStream(in).initialize();
            List<FieldDefinition> formatFileKeys = getFormatFileKeys(formatter);
            if (!formatFileKeys.isEmpty()) {
            EXCEL_KEY_LOOP:
                for (String excelKey : responseMap.keySet()) {
                    for (FieldDefinition fd : formatFileKeys) {
                        if (fd.getName().equals(excelKey)) {
                            continue EXCEL_KEY_LOOP;
                        }
                    }
                    throw new InvalidDataFormatException("応答電文フォーマット定義ファイルに存在しない業務項目が"
                            + "応答電文用スタブデータ(Excel)に存在しています。"
                            + "ファイル名:[" + new File(settings.stubResponseBookPath
                                    + settings.stubResponseBookName).getAbsolutePath() + "] "
                            + "シート名:[" + settings.stubResponseSheetName + "] 電文ID:[" + telegramId + "] "
                            + "項目名:[" + excelKey + "]");
                }
            }
        } catch (IOException e) {
            throw new InvalidDataFormatException("応答電文フォーマット定義ファイルの構文に誤りがあります。", e);
        }
    }

    /**
     * @param formatter フォーマッタ
     * @return 元になる空応答電文
     */
    private static List<FieldDefinition> getFormatFileKeys(DataRecordFormatter formatter) {

        try {
            Field filed = DataRecordFormatterSupport.class.getDeclaredField("definition");
            filed.setAccessible(true);
            LayoutDefinition def = (LayoutDefinition) filed.get(formatter);
            return def.getRecords().get(0).getFields();

        } catch (ReflectiveOperationException e) {
            // NOP
        }
        return Collections.emptyList();
    }

    /**
     * コンポーネントの戻り値（デフォルトの応答電文Map）を取得する。
     * デフォルトデータに該当する電文IDデータが存在しない場合は、警告ログを出力する。
     *
     * @param telegramId 戻り値に紐付く電文ID
     * @param requestMap 要求電文Map
     * @return 電文IDに紐付く戻り値
     */
    private static Map<String, Object> createDefaultResponseMap(String telegramId, Map<String, Object> requestMap) {

        // 電文IDに紐付くデフォルトデータを取得
        Map<String, Map<String, Object>> telegramIdMap = defaultResponseMap.get(telegramId);

        Map<String, Object> resultMap = new HashMap<>();
        if (telegramIdMap == null || telegramIdMap.isEmpty()) {
            LOGGER.logWarn(String.format("no default data exists. book:[%s]. sheet:[%s]. telegramId:[%s].",
                    settings.defaultResponseBookPath + settings.defaultResponseBookName,
                    settings.defaultResponseSheetName, telegramId));
        } else {
            // デフォルトデータは1電文に1件を前提として固定で先頭を取得
            Map<String, Object> responseDataMap = telegramIdMap.values().iterator().next();

            // デフォルトデータをベースとして、要求電文Mapに存在する値で上書きする
            resultMap.putAll(responseDataMap);
        }
        resultMap.putAll(requestMap);
        return resultMap;
    }

    /**
     * 項目値有無判定。
     * @param o スタブデータのセル項目値
     * @return 項目値が存在する場合はtrue、項目値が存在しない場合はfalse
     */
    private static boolean hasValue(Object o) {
        if (o instanceof String) {
            return StringUtil.hasValue((String) o);
        }
        return o != null;
    }

    /**
     * 応答電文Mapが例外用データの場合、例外を送出する。
     * @param map 応答電文Map
     * @return IllegalStateException 例外
     */
    private static IllegalStateException checkExceptionCase(Map<String, Object> map) {
        final String raiseExceptionKey = "raiseException"; // 例外送出判定キー
        final String raiseExceptionValue = "Exception"; // 例外送出判定値

        // 例外送出判定キーが応答電文Mapに含まれたままの場合、BeanUtilで例外スタックが出力されるためremoveで取得
        String exceptionSetting = (String) map.remove(raiseExceptionKey);
        if (raiseExceptionValue.equalsIgnoreCase(exceptionSetting)) {
            return new IllegalStateException("raiseException case:" + exceptionSetting);
        }
        return null;
    }

    /**
     * 当該電文IDのキー項目名リストとレスポンスマップからキー値の文字列を生成する。
     * キー項目が複数存在する場合は、TABをデリミタとして結合することで一意性を保持する。
     *
     * @param keyList キー項目名リスト
     * @param targetMap キー生成の対象マップ
     * @return キー値の文字列
     */
    private static String createKey(List<String> keyList, Map<String, Object> targetMap) {

        // 全てのキー項目の値を取得
        List<String> resultList = new ArrayList<>(keyList.size());
        for (String key : keyList) {
            resultList.add(String.valueOf(targetMap.get(key)));
        }

        // TABをデリミタとして結合
        return String.join("\t", resultList.toArray(new String[resultList.size()]));
    }

    /**
     * セルの値を文字列として取得する。
     * 空セルの場合は空文字を返却する。
     * @param cell セル
     * @return セルの文字列値
     */
    private static String getStringCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        // 数式は計算結果を取得
        if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
            return getStringCellValue(
                    cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator().evaluateInCell(cell));
        }
        // 文字列として取得
        cell.setCellType(Cell.CELL_TYPE_STRING);
        String result = cell.getStringCellValue();
        return result == null ? "" : result;
    }

    /**
     * settingsを設定する。
     * @param settings settings
     */
    public static void setSettings(TcpClientStubSettings settings) {
        TcpClientStub.settings = settings;
    }

    /**
     * フォーマッタファクトリを設定する。
     * @param formatterFactory フォーマッタファクトリ
     */
    public static void setFormatterFactory(FormatterFactory formatterFactory) {
        TcpClientStub.formatterFactory = formatterFactory;
    }

    /**
     * TCPクライアントスタブ設定
     * @since 1.0
     */
    public static class TcpClientStubSettings {

        /** UTリソース読み込み時のベースディレクトリ */
        private String utDefaultResourcesRoot = SystemRepository.get("N.test.resource-root");

        /** デフォルトデータ用ブックパス */
        private String defaultResponseBookPath = "src/test/resources/jp/S/common/tcp/default/";

        /** デフォルトデータ用ブック名 */
        private String defaultResponseBookName = "defaultSocketResponse.xlsx";

        /** デフォルトデータ用シート名 */
        private String defaultResponseSheetName = "SocketResonse";

        /** スタブデータ用ブックパス */
        private String stubResponseBookPath = "src/test/resources/jp/S/common/tcp/";

        /** スタブデータ用ブック名 */
        private String stubResponseBookName = ""; // socketResponse.xlsx

        /** スタブデータ用シート名 */
        private String stubResponseSheetName = "SocketResonse";

        /**
         * UTリソース読み込み時のベースディレクトリを設定する。
         * @param utDefaultResourcesRoot UTリソース読み込み時のベースディレクトリ
         */
        public void setUtDefaultResourcesRoot(String utDefaultResourcesRoot) {
            if (StringUtil.hasValue(utDefaultResourcesRoot)) {
                this.utDefaultResourcesRoot = utDefaultResourcesRoot;
            }
        }

        /**
         * デフォルトデータ用ブックパスを設定する。
         * @param defaultResponseBookPath デフォルトデータ用ブックパス
         */
        public void setDefaultResponseBookPath(String defaultResponseBookPath) {
            if (StringUtil.hasValue(defaultResponseBookPath)) {
                this.defaultResponseBookPath = defaultResponseBookPath;
            }
        }

        /**
         * デフォルトデータ用ブックパスを設定する。
         * @param defaultResponseBookName デフォルトデータ用ブックパス
         */
        public void setDefaultResponseBookName(String defaultResponseBookName) {
            if (StringUtil.hasValue(defaultResponseBookName)) {
                this.defaultResponseBookName = defaultResponseBookName;
            }
        }

        /**
         * デフォルトデータ用シート名を設定する。
         * @param defaultResponseSheetName デフォルトデータ用シート名
         */
        public void setDefaultResponseSheetName(String defaultResponseSheetName) {
            if (StringUtil.hasValue(defaultResponseSheetName)) {
                this.defaultResponseSheetName = defaultResponseSheetName;
            }
        }

        /**
         * スタブデータ用ブックパスを設定する。
         * @param stubResponseBookPath スタブデータ用ブックパス
         */
        public void setStubResponseBookPath(String stubResponseBookPath) {
            if (StringUtil.hasValue(stubResponseBookPath)) {
                this.stubResponseBookPath = stubResponseBookPath;
            }
        }

        /**
         * スタブデータ用ブック名を設定する。
         * @param stubResponseBookName スタブデータ用ブック名
         */
        public void setStubResponseBookName(String stubResponseBookName) {
            if (StringUtil.hasValue(stubResponseBookName)) {
                this.stubResponseBookName = stubResponseBookName;
            }
        }

        /**
         * スタブデータ用シート名を設定する。
         * @param stubResponseSheetName スタブデータ用シート名
         */
        public void setStubResponseSheetName(String stubResponseSheetName) {
            if (StringUtil.hasValue(stubResponseSheetName)) {
                this.stubResponseSheetName = stubResponseSheetName;
            }
        }
    }
}
