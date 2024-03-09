package top.kongsheng.common.word.datainput.model.input;

import cn.hutool.core.io.IoUtil;
import cn.hutool.core.lang.Pair;
import top.kongsheng.common.word.datainput.model.FullContent;
import top.kongsheng.common.word.datainput.utils.FileBase64Util;
import org.apache.poi.common.usermodel.PictureType;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;

/**
 * 图片填充类
 *
 * @author 孔胜
 * @date 2023/8/3 10:56
 */
public class ImageDataInput extends AbsDataInput<ImageDataInput> {

    /**
     * 字节数组
     */
    private byte[] bytes;

    private PictureType pictureType;

    protected ImageDataInput() {
        super(2);
    }

    public byte[] getBytes() {
        return bytes;
    }

    public void setBytes(byte[] bytes) {
        this.bytes = bytes;
    }

    public PictureType getPictureType() {
        return pictureType;
    }

    public void setPictureType(PictureType pictureType) {
        this.pictureType = pictureType;
    }

    @Override
    public void full(FullContent fullContent) {
        XWPFRun run = fullContent.getRun();
        double width = fullContent.getWidth();
        try (
                ByteArrayInputStream inputStream = IoUtil.toStream(bytes);
                ByteArrayOutputStream outputStream = new ByteArrayOutputStream()
        ) {
            String extension = pictureType.getExtension();
            BufferedImage bufferedImage = ImageIO.read(inputStream);
            float sourceWidth = bufferedImage.getWidth();
            float sourceHeight = bufferedImage.getHeight();
            float aspectRatio = sourceWidth / sourceHeight;
            float targetWidth = (float) (width * 0.95F);
            int intWidth = (int) Math.rint(targetWidth * Units.EMU_PER_DXA);
            int intHeight = (int) Math.rint((targetWidth / aspectRatio) * Units.EMU_PER_DXA);
            ImageIO.write(bufferedImage, extension.substring(1), outputStream);
            try (ByteArrayInputStream pictureData = IoUtil.toStream(outputStream)) {
                run.addPicture(pictureData, pictureType, "image" + extension, intWidth, intHeight);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static ImageDataInput base64To(String base64) {
        ImageDataInput imageDataInput = new ImageDataInput();
        Pair<String, String> fileBase64 = FileBase64Util.handleFileBase64(base64);
        imageDataInput.setBytes(FileBase64Util.toBytes(fileBase64.getKey()));
        PictureType pictureType;
        switch (fileBase64.getValue()) {
            case "png":
                pictureType = PictureType.PNG;
                break;
            case "jpg":
            default:
                pictureType = PictureType.JPEG;
                break;
        }
        imageDataInput.setPictureType(pictureType);
        return imageDataInput;
    }

    @Override
    public long dataLength() {
        if (bytes == null) {
            return 0;
        }
        return bytes.length;
    }

    @Override
    public boolean isEmpty() {
        return bytes == null || bytes.length < 1;
    }
}
