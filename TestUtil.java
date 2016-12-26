import java.lang.reflect.Field;

/**
 *
 */
public class TestUtil {

    /** private constructor */
    private TestUtil() {
        // NOP
    }

    /**
     * 対象オブジェクトの内部フィールド値を取得する。
     *
     * @param <T> 内部フィールドクラス
     * @param targetClass 対象クラス
     * @param target 対象オブジェクト
     * @param fieldName フィールド名
     * @return フィールド値
     */
    @SuppressWarnings("unchecked")
    public static <T> T getFieldObject(Class<?> targetClass, Object target, String fieldName) {
        try {
            Field field = targetClass.getDeclaredField(fieldName);
            field.setAccessible(true);
            return (T) field.get(target);
        } catch (ReflectiveOperationException e) {
            e.printStackTrace();
            return null;
        }
    }

    /**
     * 対象オブジェクトの内部フィールド値を設定する。
     *
     * @param target 対象オブジェクト
     * @param fieldName フィールド名
     * @param value フィールド値
     */
    public static void setFieldObject(Object target, String fieldName, Object value) {
        try {
            Field field = target.getClass().getDeclaredField(fieldName);
            field.setAccessible(true);
            field.set(target, value);
        } catch (ReflectiveOperationException e) {
            e.printStackTrace();
        }
    }

}
