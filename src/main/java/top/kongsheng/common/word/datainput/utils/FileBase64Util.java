package top.kongsheng.common.word.datainput.utils;

import cn.hutool.core.lang.Pair;
import cn.hutool.core.util.StrUtil;

import java.util.Base64;

/**
 * 文件Base64工具类
 *
 * @author 孔胜
 * @date 2023/8/1 20:01
 */
public class FileBase64Util {

    public static byte[] toBytes(String fileBase64) {
        byte[] fileByte64 = Base64.getDecoder().decode(fileBase64);
        return fileByte64;
    }


    public static Pair<String, String> handleFileBase64(String fileBase64) {
        if (StrUtil.isEmpty(fileBase64)) {
            throw new RuntimeException("转换失败，Base64值为空。");
        }
        int index = fileBase64.indexOf(";");
        if (index < 0) {
            throw new RuntimeException("转换失败，无效的Base64。");
        }
        String dataInfo = fileBase64.substring(0, index);
        String fileTypeInfo = dataInfo.replace("data:", "");
        String[] fileTypeInfos = fileTypeInfo.split("/");
        String fileType = fileTypeInfos[0];
        String suffix;
        if ("image".equals(fileType)) {
            String fileSuffix = fileTypeInfos[1];
            switch (fileSuffix) {
                case "jpeg":
                    suffix = "jpg";
                    break;
                case "png":
                    suffix = "png";
                    break;
                case "webp":
                    suffix = "webp";
                    break;
                case "gif":
                case "svg+xml":
                default:
                    throw new RuntimeException("转换失败，暂不支持(" + fileTypeInfo + ")类型的Base64进行转换。");
            }
        } else {
            throw new RuntimeException("转换失败，暂不支持(" + fileTypeInfo + ")类型的Base64进行转换。");
        }
        String base = fileBase64.substring(index + 1 + "base64,".length());
        return new Pair<>(base, suffix);
    }
}
