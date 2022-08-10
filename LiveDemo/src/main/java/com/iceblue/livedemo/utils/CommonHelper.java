package com.iceblue.livedemo.utils;

import org.dom4j.Element;
import org.dom4j.io.SAXReader;

import javax.activation.MimetypesFileTypeMap;
import javax.imageio.ImageIO;
import javax.servlet.http.HttpServletResponse;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.lang.reflect.Method;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.UUID;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import static org.springframework.util.StreamUtils.BUFFER_SIZE;

public class CommonHelper {
    /**
     * 从Xml读取对象数据 直接映射到对象
     * 注意：
     * 1. xml结构必须是两层，在根节点内就应该是model的信息。否则需要调整代码解析逻辑。例如下面:
     * <root>
     * <Person>
     * <Name>张三</Name>
     * <Age>25</Age>
     * </Person>
     * <Person>
     * <Name>李四</Name>
     * <Age>26</Age>
     * </Person>
     * ....
     * </root>
     * 2. xml中的属性名字必须与实体model里的set方法后缀一样（因为需要反射调用set方法）。例如，
     * Person类中的属性name，对应的set方法就是setName，所以xml里这个属性必须为首字母大写的Name
     * <p>
     * 3. 必须保留实体model中的默认构造方法，即不传入参数的构造方法
     *
     * @param xmlFilePath xml文件名
     * @param clazz       需要映射的model的class
     * @param <T>         需要映射的model类型
     * @return 返回读取到的model列表
     */
    public static <T> List<T> xmlToBean(String xmlFilePath, Class<T> clazz) {
        List<T> list = new ArrayList<>();
        SAXReader reader = new SAXReader();
        org.dom4j.Document document = null;
        try {
            document = reader.read(new File(xmlFilePath));
            Element rootElement = document.getRootElement();
            Iterator iterator = rootElement.elementIterator();
            while (iterator.hasNext()) {
                Element country = (Element) iterator.next();
                T model = clazz.newInstance();
                Iterator attrsIterator = country.elementIterator();
                while (attrsIterator.hasNext()) {
                    Element couAttr = (Element) attrsIterator.next();
                    String fieldName = couAttr.getName();
                    //getDeclaredField用于获取private修饰的属性
                    Class<?> filedType = clazz.getDeclaredField(lowerOrUpper(fieldName)).getType();
                    Class<?>[] params = {filedType};
                    //通过反射调用model的set方法，为model的各个属性设置值
                    Reflection(model, clazz, "set" + couAttr.getName(), params, castArgsType(filedType, couAttr.getStringValue()));
                }
                list.add(model);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return list;
    }

    /**
     * 转换参数类型
     *
     * @param argClass
     * @param argStr
     * @return
     */
    public static Object castArgsType(Class<?> argClass, String argStr) {
        switch (argClass.getTypeName()) {
            case "int":
                return Integer.parseInt(argStr);
            case "long":
                return Long.parseLong(argStr);
            case "double":
                return Double.parseDouble(argStr);
            case "Float":
                return Float.parseFloat(argStr);
            default:
                return argStr;
        }
    }

    /**
     * 首字母大小写转换
     *
     * @param oldStr
     * @return
     */
    public static String lowerOrUpper(String oldStr) {
        char[] chars = oldStr.toCharArray();
        if (chars[0] >= 65 && chars[0] <= 90) {
            chars[0] += 32;
        } else if (chars[0] >= 97 && chars[0] <= 122) {
            chars[0] -= 32;
        }
        return String.valueOf(chars);

    }

    /**
     * 通过反射执行clazz的方法
     *
     * @param object     ---调到该方法的具体对象
     * @param clazz      ---具体对象的class类型
     * @param params     ---反射方法中参数class类型
     * @param methodName ---反射方法的名称
     * @param args       ----调用该方法用到的具体参数
     */
    public static Object Reflection(Object object, Class<?> clazz, String methodName, Class<?>[] params, Object... args) {
        try {
            Method method = clazz.getMethod(methodName, params);
            return method.invoke(object, args);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }
    /*
    public static List<TypeOptionsModel> getFontOptions() {
        List<TypeOptionsModel> typeOptionsModelList = new ArrayList<>();
        typeOptionsModelList.add(new TypeOptionsModel("Aharoni", "Aharoni"));
        typeOptionsModelList.add(new TypeOptionsModel("Andalus", "Andalus"));
        typeOptionsModelList.add(new TypeOptionsModel("Angsana New", "Angsana New"));
        typeOptionsModelList.add(new TypeOptionsModel("AngsanaUPC", "AngsanaUPC"));
        typeOptionsModelList.add(new TypeOptionsModel("Aparajita", "Aparajita"));
        typeOptionsModelList.add(new TypeOptionsModel("Arabic Typesetting", "Arabic Typesetting"));
        typeOptionsModelList.add(new TypeOptionsModel("Arial", "Arial"));
        typeOptionsModelList.add(new TypeOptionsModel("Arial Black", "Arial Black"));
        typeOptionsModelList.add(new TypeOptionsModel("Arial Unicode MS", "Arial Unicode MS"));
        typeOptionsModelList.add(new TypeOptionsModel("Batang", "Batang"));
        typeOptionsModelList.add(new TypeOptionsModel("BatangChe", "BatangChe"));
        typeOptionsModelList.add(new TypeOptionsModel("Browallia New", "Browallia New"));
        typeOptionsModelList.add(new TypeOptionsModel("BrowalliaUPC", "BrowalliaUPC"));
        typeOptionsModelList.add(new TypeOptionsModel("Calibri", "Calibri"));
        typeOptionsModelList.add(new TypeOptionsModel("Calibri Light", "Calibri Light"));
        typeOptionsModelList.add(new TypeOptionsModel("Cambria", "Cambria"));
        typeOptionsModelList.add(new TypeOptionsModel("Cambria Math", "Cambria Math"));
        typeOptionsModelList.add(new TypeOptionsModel("Comic Sans MS", "Comic Sans MS"));
        typeOptionsModelList.add(new TypeOptionsModel("Consolas", "Consolas"));
        typeOptionsModelList.add(new TypeOptionsModel("Constantia", "Constantia"));
        typeOptionsModelList.add(new TypeOptionsModel("Corbel", "Corbel"));
        typeOptionsModelList.add(new TypeOptionsModel("Cordia New", "Cordia New"));
        typeOptionsModelList.add(new TypeOptionsModel("CordiaUPC", "CordiaUPC"));
        typeOptionsModelList.add(new TypeOptionsModel("Courier New", "Courier New"));
        typeOptionsModelList.add(new TypeOptionsModel("DaunPenh", "DaunPenh"));
        typeOptionsModelList.add(new TypeOptionsModel("David", "David"));
        typeOptionsModelList.add(new TypeOptionsModel("DFKai-SB", "DFKai-SB"));
        typeOptionsModelList.add(new TypeOptionsModel("DilleniaUPC", "DilleniaUPC"));
        typeOptionsModelList.add(new TypeOptionsModel("DokChampa", "DokChampa"));
        typeOptionsModelList.add(new TypeOptionsModel("Dotum", "Dotum"));
        typeOptionsModelList.add(new TypeOptionsModel("DotumChe", "DotumChe"));
        typeOptionsModelList.add(new TypeOptionsModel("Ebrima", "Ebrima"));
        typeOptionsModelList.add(new TypeOptionsModel("Estrangelo Edessa", "Estrangelo Edessa"));
        typeOptionsModelList.add(new TypeOptionsModel("Euphemia", "Euphemia"));
        typeOptionsModelList.add(new TypeOptionsModel("FangSong", "FangSong"));
        typeOptionsModelList.add(new TypeOptionsModel("Franklin Gothic Medium", "Franklin Gothic Medium"));
        typeOptionsModelList.add(new TypeOptionsModel("FrankRuehl", "FrankRuehl"));
        typeOptionsModelList.add(new TypeOptionsModel("FreesiaUPC", "FreesiaUPC"));
        typeOptionsModelList.add(new TypeOptionsModel("Gabriola", "Gabriola"));
        typeOptionsModelList.add(new TypeOptionsModel("Gautami", "Gautami"));
        typeOptionsModelList.add(new TypeOptionsModel("Georgia", "Georgia"));
        typeOptionsModelList.add(new TypeOptionsModel("Gisha", "Gisha"));
        typeOptionsModelList.add(new TypeOptionsModel("Gulim", "Gulim"));
        typeOptionsModelList.add(new TypeOptionsModel("GulimChe", "GulimChe"));
        typeOptionsModelList.add(new TypeOptionsModel("Gungsuh", "Gungsuh"));
        typeOptionsModelList.add(new TypeOptionsModel("GungsuhChe", "GungsuhChe"));
        typeOptionsModelList.add(new TypeOptionsModel("Impact", "Impact"));
        typeOptionsModelList.add(new TypeOptionsModel("lrisUPC", "lrisUPC"));
        typeOptionsModelList.add(new TypeOptionsModel("Iskoola Pota", "Iskoola Pota"));
        typeOptionsModelList.add(new TypeOptionsModel("JasmineUPC", "JasmineUPC"));
        typeOptionsModelList.add(new TypeOptionsModel("KaiTi", "KaiTi"));
        typeOptionsModelList.add(new TypeOptionsModel("Kalinga", "Kalinga"));
        typeOptionsModelList.add(new TypeOptionsModel("Kartika", "Kartika"));
        typeOptionsModelList.add(new TypeOptionsModel("Khmer UI", "Khmer UI"));
        typeOptionsModelList.add(new TypeOptionsModel("KodchiangUPC", "KodchiangUPC"));
        typeOptionsModelList.add(new TypeOptionsModel("Kokila", "Kokila"));
        typeOptionsModelList.add(new TypeOptionsModel("Lao UI", "Lao UI"));
        typeOptionsModelList.add(new TypeOptionsModel("Latha", "Latha"));
        typeOptionsModelList.add(new TypeOptionsModel("Leelawadee", "Leelawadee"));
        typeOptionsModelList.add(new TypeOptionsModel("LilyUPC", "LilyUPC"));
        typeOptionsModelList.add(new TypeOptionsModel("Lucida Console", "Lucida Console"));
        typeOptionsModelList.add(new TypeOptionsModel("Lucida Sans Unicode", "Lucida Sans Unicode"));
        typeOptionsModelList.add(new TypeOptionsModel("Malgun Gothic", "Malgun Gothic"));
        typeOptionsModelList.add(new TypeOptionsModel("Mangal", "Mangal"));
        typeOptionsModelList.add(new TypeOptionsModel("Marlett", "Marlett"));
        typeOptionsModelList.add(new TypeOptionsModel("Meiryo", "Meiryo"));
        typeOptionsModelList.add(new TypeOptionsModel("Meiryo UI", "Meiryo UI"));
        typeOptionsModelList.add(new TypeOptionsModel("Microsoft Himalaya", "Microsoft Himalaya"));
        typeOptionsModelList.add(new TypeOptionsModel("Microsoft JhengHei", "Microsoft JhengHei"));
        typeOptionsModelList.add(new TypeOptionsModel("Microsoft New Tai Lue", "Microsoft New Tai Lue"));
        typeOptionsModelList.add(new TypeOptionsModel("Microsoft PhagsPa", "Microsoft PhagsPa"));
        typeOptionsModelList.add(new TypeOptionsModel("Microsoft Sans Serif", "Microsoft Sans Serif"));
        typeOptionsModelList.add(new TypeOptionsModel("Microsoft Tai Le", "Microsoft Tai Le"));
        typeOptionsModelList.add(new TypeOptionsModel("Microsoft Uighur", "Microsoft Uighur"));
        typeOptionsModelList.add(new TypeOptionsModel("Microsoft YaHei", "Microsoft YaHei"));
        typeOptionsModelList.add(new TypeOptionsModel("Microsoft Yi Baiti", "Microsoft Yi Baiti"));
        typeOptionsModelList.add(new TypeOptionsModel("MingLiU", "MingLiU"));
        typeOptionsModelList.add(new TypeOptionsModel("MingLiU-ExtB", "MingLiU-ExtB"));
        typeOptionsModelList.add(new TypeOptionsModel("MingLiU_HKSCS", "MingLiU_HKSCS"));
        typeOptionsModelList.add(new TypeOptionsModel("MingLiU_HKSCS-ExtB", "MingLiU_HKSCS-ExtB"));
        typeOptionsModelList.add(new TypeOptionsModel("Miriam", "Miriam"));
        typeOptionsModelList.add(new TypeOptionsModel("Miriam Fixed", "Miriam Fixed"));
        typeOptionsModelList.add(new TypeOptionsModel("Mongolian Baiti", "Mongolian Baiti"));
        typeOptionsModelList.add(new TypeOptionsModel("MoolBoran", "MoolBoran"));
        typeOptionsModelList.add(new TypeOptionsModel("MS Gothic", "MS Gothic"));
        typeOptionsModelList.add(new TypeOptionsModel("MS Mincho", "MS Mincho"));
        typeOptionsModelList.add(new TypeOptionsModel("MS PGothic", "MS PGothic"));
        typeOptionsModelList.add(new TypeOptionsModel("MS PMincho", "MS PMincho"));
        typeOptionsModelList.add(new TypeOptionsModel("MS UI Gothic", "MS UI Gothic"));
        typeOptionsModelList.add(new TypeOptionsModel("MV Boli", "MV Boli"));
        typeOptionsModelList.add(new TypeOptionsModel("Narkisim", "Narkisim"));
        typeOptionsModelList.add(new TypeOptionsModel("NSimSun", "NSimSun"));
        typeOptionsModelList.add(new TypeOptionsModel("Nyala", "Nyala"));
        typeOptionsModelList.add(new TypeOptionsModel("Palatino Linotype", "Palatino Linotype"));
        typeOptionsModelList.add(new TypeOptionsModel("Plantagenet Cherokee", "Plantagenet Cherokee"));
        typeOptionsModelList.add(new TypeOptionsModel("PMingLiU", "PMingLiU"));
        typeOptionsModelList.add(new TypeOptionsModel("PMingLiU-ExtB", "PMingLiU-ExtB"));
        typeOptionsModelList.add(new TypeOptionsModel("Raavi", "Raavi"));
        typeOptionsModelList.add(new TypeOptionsModel("Rod", "Rod"));
        typeOptionsModelList.add(new TypeOptionsModel("Sakkal Majalla", "Sakkal Majalla"));
        typeOptionsModelList.add(new TypeOptionsModel("Segoe Print", "Segoe Print"));
        typeOptionsModelList.add(new TypeOptionsModel("Segoe Script", "Segoe Script"));
        typeOptionsModelList.add(new TypeOptionsModel("Segoe UI", "Segoe UI"));
        typeOptionsModelList.add(new TypeOptionsModel("Segoe UI Light", "Segoe UI Light"));
        typeOptionsModelList.add(new TypeOptionsModel("Segoe UI Semibold", "Segoe UI Semibold"));
        typeOptionsModelList.add(new TypeOptionsModel("Segoe UI Symbol", "Segoe UI Symbol"));
        typeOptionsModelList.add(new TypeOptionsModel("Shonar Bangla", "Shonar Bangla"));
        typeOptionsModelList.add(new TypeOptionsModel("Shruti", "Shruti"));
        typeOptionsModelList.add(new TypeOptionsModel("SimHei", "SimHei"));
        typeOptionsModelList.add(new TypeOptionsModel("Simplified Arabic", "Simplified Arabic"));
        typeOptionsModelList.add(new TypeOptionsModel("Simplified Arabic Fixed", "Simplified Arabic Fixed"));
        typeOptionsModelList.add(new TypeOptionsModel("SimSun", "SimSun"));
        typeOptionsModelList.add(new TypeOptionsModel("SimSun-ExtB", "SimSun-ExtB"));
        typeOptionsModelList.add(new TypeOptionsModel("Sylfaen", "Sylfaen"));
        typeOptionsModelList.add(new TypeOptionsModel("Symbol", "Symbol"));
        typeOptionsModelList.add(new TypeOptionsModel("Tahoma", "Tahoma"));
        typeOptionsModelList.add(new TypeOptionsModel("Times New Roman", "Times New Roman"));
        typeOptionsModelList.add(new TypeOptionsModel("Traditional Arabic", "Traditional Arabic"));
        typeOptionsModelList.add(new TypeOptionsModel("Trebuchet MS", "Trebuchet MS"));
        typeOptionsModelList.add(new TypeOptionsModel("Tunga", "Tunga"));
        typeOptionsModelList.add(new TypeOptionsModel("Utsaah", "Utsaah"));
        typeOptionsModelList.add(new TypeOptionsModel("Vani", "Vani"));
        typeOptionsModelList.add(new TypeOptionsModel("Verdana", "Verdana"));
        typeOptionsModelList.add(new TypeOptionsModel("Vijaya", "Vijaya"));
        typeOptionsModelList.add(new TypeOptionsModel("Vrinda", "Vrinda"));
        typeOptionsModelList.add(new TypeOptionsModel("Webdings", "Webdings"));
        typeOptionsModelList.add(new TypeOptionsModel("Wingdings", "Wingdings"));


        return typeOptionsModelList;
    }*/

    /**
     * @param fileName：
     * @return
     */
    public static boolean delete(String fileName) {
        File file = new File(fileName);
        if (!file.exists()) {
            return false;
        } else {
            if (file.isFile())
                return deleteFile(fileName);
            else
                return deleteDirectory(fileName);
        }
    }

    /**
     * @param fileName：
     * @return
     */
    public static boolean deleteFile(String fileName) {
        File file = new File(fileName);
        if (file.exists() && file.isFile()) {
            if (file.delete()) {
                return true;
            } else {
                return false;
            }
        } else {
            return false;
        }
    }

    /**
     * @param dir：
     * @return
     */
    public static boolean deleteDirectory(String dir) {
        if (!dir.endsWith(File.separator))
            dir = dir + File.separator;
        File dirFile = new File(dir);
        if ((!dirFile.exists()) || (!dirFile.isDirectory())) {
            return false;
        }
        boolean flag = true;
        File[] files = dirFile.listFiles();
        for (int i = 0; i < files.length; i++) {
            if (files[i].isFile()) {
                flag = deleteFile(files[i].getAbsolutePath());
                if (!flag)
                    break;
            } else if (files[i].isDirectory()) {
                flag = deleteDirectory(files[i].getAbsolutePath());
                if (!flag)
                    break;
            }
        }
        if (!flag) {
            return false;
        }
        if (dirFile.delete()) {
            return true;
        } else {
            return false;
        }
    }

    /**
     * @param srcDir           压缩文件夹路径
     * @param out              压缩文件输出流
     * @param KeepDirStructure 是否保留原来的目录结构,true:保留目录结构;
     *                         false:所有文件跑到压缩包根目录下(注意：不保留目录结构可能会出现同名文件,会压缩失败)
     * @throws RuntimeException 压缩失败会抛出运行时异常
     */
    public static void toZip(String srcDir, OutputStream out, boolean KeepDirStructure)
            throws RuntimeException, IOException {
        ZipOutputStream zos = null;
        try {
            zos = new ZipOutputStream(out);
            File sourceFile = new File(srcDir);
            compress(sourceFile, zos, sourceFile.getName(), KeepDirStructure);
        } catch (Exception e) {
            throw new RuntimeException("zip error from ZipUtils", e);
        } finally {
            if (zos != null) {
                zos.close();
            }
            if (out != null) {
                out.close();
            }
        }
    }

    /**
     * 递归压缩
     *
     * @param sourceFile
     * @param zos
     * @param name
     * @param KeepDirStructure
     * @throws Exception
     */
    private static void compress(File sourceFile, ZipOutputStream zos, String name,
                                 boolean KeepDirStructure) throws Exception {
        byte[] buf = new byte[BUFFER_SIZE];
        if (sourceFile.isFile()) {
            // 向zip输出流中添加一个zip实体，构造器中name为zip实体的文件的名字
            zos.putNextEntry(new ZipEntry(name));
            // copy文件到zip输出流中
            int len;
            FileInputStream in = new FileInputStream(sourceFile);
            while ((len = in.read(buf)) != -1) {
                zos.write(buf, 0, len);
            }
            // Complete the entry
            zos.closeEntry();
            in.close();
        } else {
            File[] listFiles = sourceFile.listFiles();
            if (listFiles == null || listFiles.length == 0) {
                // 需要保留原来的文件结构时,需要对空文件夹进行处理
                if (KeepDirStructure) {
                    // 空文件夹的处理
                    zos.putNextEntry(new ZipEntry(name + "/"));
                    // 没有文件，不需要文件的copy
                    zos.closeEntry();
                }
            } else {
                for (File file : listFiles) {
                    // 判断是否需要保留原来的文件结构
                    if (KeepDirStructure) {
                        // 注意：file.getName()前面需要带上父文件夹的名字加一斜杠,
                        // 不然最后压缩包中就不能保留原来的文件结构,即：所有文件都跑到压缩包根目录下了
                        compress(file, zos, name + "/" + file.getName(), true);
                    } else {
                        compress(file, zos, file.getName(), false);
                    }
                }
            }
        }
    }

    /**
     * 打包图片
     *
     * @param images
     * @return
     */
    public static String packageImages(BufferedImage[] images, String type) throws IOException {
        String zipFileName = Static.OUTPUT_FILE_START + UUID.randomUUID() + ".zip";
        File zipFile = new File(Static.OUTPUT_FILE_PATH + zipFileName);
        ZipOutputStream zos = null;
        FileOutputStream out = null;
        try {
            out = new FileOutputStream(zipFile);
            zos = new ZipOutputStream(out);
            int index = 1;
            for (BufferedImage image : images) {
                byte[] buf = imageToBytes(image, type);
                zos.putNextEntry(new ZipEntry(Static.OUTPUT_IMAGE_START + index + "." + type));
                zos.write(buf, 0, buf.length);
                zos.closeEntry();
                index++;
            }
        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        } finally {
            if (zos != null) {
                zos.close();
            }
            if (out != null) {
                out.close();
            }
        }
        return zipFileName;
    }

    /**
     * image转byte数组
     *
     * @param image
     * @param type
     * @return
     * @throws IOException
     */
    public static byte[] imageToBytes(BufferedImage image, String type) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ImageIO.write(image, type, out);
        return out.toByteArray();
    }

    /**
     * 下载文件
     *
     * @param response
     * @param fileName
     * @throws IOException
     */
    public static void downloadFile(HttpServletResponse response, String path, String fileName, boolean deleteFile) throws IOException {
        File file = new File(path + fileName);
        if (!file.exists()) {
            return;
        }
        String type = new MimetypesFileTypeMap().getContentType(fileName);
        response.setHeader("Content-type", type);
        String encode = new String(fileName.getBytes(StandardCharsets.UTF_8), StandardCharsets.ISO_8859_1);
        response.setHeader("Content-Disposition", "attachment;filename=" + encode);

        OutputStream outputStream = response.getOutputStream();
        byte[] buff = new byte[1024];
        BufferedInputStream inputStream = null;
        inputStream = new BufferedInputStream(new FileInputStream(file));
        int len=0;
        while ( (len=inputStream.read(buff))>0) {
            outputStream.write(buff, 0,len);
        }
        inputStream.close();

        if (deleteFile) {
            delete(path + fileName);
        }
    }

    /**
     * 字符串转Color
     * 例如 "#000000"为黑色
     *
     * @param str
     * @return
     */
    public static Color String2Color(String str) {
        try {
            int rgb = Integer.parseInt(str, 16);
            Color color = null;
            color = new Color(rgb);
            return color;
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }
}

