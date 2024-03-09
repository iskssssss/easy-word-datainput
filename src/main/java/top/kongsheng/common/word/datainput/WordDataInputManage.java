package top.kongsheng.common.word.datainput;

import cn.hutool.core.io.FileUtil;
import top.kongsheng.common.word.datainput.convert.InputDataConvert;
import top.kongsheng.common.word.datainput.core.WordDataInputCore;
import top.kongsheng.common.word.datainput.exception.FullException;
import top.kongsheng.common.word.datainput.model.TemplateInfo;

import java.io.File;
import java.util.List;

/**
 * word数据填充管理类
 *
 * @author 孔胜
 * @date 2023/8/7 14:47
 */
public class WordDataInputManage {
    public static String WORD_DATA_INPUT_FILE_SAVE_DIR = FileUtil.getTmpDir().getPath() + File.separator + "wordDataInputFile";

    static {
        FileUtil.mkdir(WORD_DATA_INPUT_FILE_SAVE_DIR);
    }

    public static void setWordFataInputFileSaveDir(String dir) {
        WORD_DATA_INPUT_FILE_SAVE_DIR = dir;
        FileUtil.mkdir(WORD_DATA_INPUT_FILE_SAVE_DIR);
    }


    /**
     * 填充数据
     *
     * @param inputDataConvert 填充数据转换类
     * @return 结果文件
     * @throws FullException 填充一场
     */
    public static File full(TemplateInfo templateInfo, InputDataConvert inputDataConvert) throws FullException {
        try (WordDataInputCore wordDataInputCore = new WordDataInputCore(templateInfo)) {
            wordDataInputCore.setInputDataConvert(inputDataConvert);
            wordDataInputCore.run();
            File printFile = wordDataInputCore.save();
            if (printFile == null) {
                throw new FullException("文件保存失败。");
            }
            return printFile;
        }
    }

    /**
     * 创建指定后缀的文件
     *
     * @param suffix 后缀
     * @return 文件
     */
    public static File touch(String suffix) {
        String fileFullPath = WORD_DATA_INPUT_FILE_SAVE_DIR + File.separator + "SPLICE_FILE_" + System.currentTimeMillis() + "." + suffix;
        return FileUtil.touch(fileFullPath);
    }

    /**
     * 获取缓存文件列表
     *
     * @return 缓存文件列表
     */
    public static List<String> listCacheFile() {
        return FileUtil.listFileNames(WORD_DATA_INPUT_FILE_SAVE_DIR);
    }

    /**
     * 删除文件
     *
     * @param fileName 文件全名
     * @return 结果
     */
    public static boolean delCacheFile(String fileName) {
        return FileUtil.del(WORD_DATA_INPUT_FILE_SAVE_DIR + File.separator + fileName);
    }

    /**
     * 删除全部缓存文件
     *
     * @return 结果
     */
    public static boolean delAllCacheFile() {
        List<String> listFileNames = FileUtil.listFileNames(WORD_DATA_INPUT_FILE_SAVE_DIR);
        listFileNames.forEach(WordDataInputManage::delCacheFile);
        return true;
    }
}
