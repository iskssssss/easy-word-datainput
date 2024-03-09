package top.kongsheng.common.word.datainput.utils;

/**
 * 字符工具类
 *
 * @author 孔胜
 * @date 2023/9/4 16:45
 */
public class CharUtil {
    private static final Character[][] UPPER_CASE_NUMBERS = new Character[][]
            {
                    {'零', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十'},
                    {' ', ' ', '十', '百', '千', '万', ' '}
            };

    public static String toUpperCaseNumber(int num) {
        if (num < 0) {
            return "";
        }
        if (num <= 10) {
            return UPPER_CASE_NUMBERS[0][num].toString();
        }
        // 11  十一
        // 20  二十
        // 21  二十一
        // 100 一百
        // 101 一百零一
        // 110 一百一十
        // 111 一百一十一
        // 十 百 千 万
        String numStr = String.valueOf(num);
        int numStrLength = numStr.length();
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < numStrLength; i++) {
            int i1 = numStr.charAt(i) - '0';
            char x = UPPER_CASE_NUMBERS[0][i1];
            char x1 = UPPER_CASE_NUMBERS[1][numStrLength - i];
            if (i1 == 0) {
                if (i < numStrLength - 1 && (numStr.charAt(i + 1) - '0') > 0) {
                    sb.append(x);
                }
                continue;
            }
            sb.append(x).append(x1);
        }
        return sb.toString();
    }

    public static long dataCharLength(String valueStr) {
        int valueStrLength = valueStr.length();
        long dataCharLength = 0;
        for (int i = 0; i < valueStrLength; i++) {
            char c = valueStr.charAt(i);
            dataCharLength = dataCharLength + ((Character.charCount(c) > 2 || isChineseCharacter(c)) ? 2 : 1);
        }
        return dataCharLength;
    }

    public static boolean isChineseCharacter(char ch) {
        Character.UnicodeBlock block = Character.UnicodeBlock.of(ch);
        return block == Character.UnicodeBlock.CJK_UNIFIED_IDEOGRAPHS
                || block == Character.UnicodeBlock.CJK_UNIFIED_IDEOGRAPHS_EXTENSION_A
                || block == Character.UnicodeBlock.CJK_UNIFIED_IDEOGRAPHS_EXTENSION_B
                || block == Character.UnicodeBlock.CJK_COMPATIBILITY_IDEOGRAPHS
                || block == Character.UnicodeBlock.CJK_COMPATIBILITY_IDEOGRAPHS_SUPPLEMENT;
    }
}
