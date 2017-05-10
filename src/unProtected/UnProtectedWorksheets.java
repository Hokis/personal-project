package unProtected;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Enumeration;
import java.util.List;
import java.util.zip.ZipOutputStream;

import org.apache.commons.compress.archivers.zip.ZipArchiveEntry;
import org.apache.commons.compress.archivers.zip.ZipFile;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;
import org.dom4j.io.XMLWriter;

/**
 * <p>直接去除指定工作簿的所有工作表的保护密码</p>
 * <p>针对07及更高版本的Excel文件（指xlsx为后缀的文件）</p>
 * <p>环境：jdk1.8，第三方jar包：dom4j-1.6.1.jar，jaxen-1.1-beta-6.jar，commons-compress-1.13.jar</p>
 * <p>Note：由于使用jdk提供的Zip实现出了问题，网上搜索一番，决定使用Apache的实现（Apache Commons Compress）</p>
 * 
 * @author Hokis
 * 
 * @since 2017-5-10 
 * 
 * @version 1.0
 *
 */
public class UnProtectedWorksheets {

	public static void main(String[] args) throws IOException {
		//目标xlsx文件
		File oldZipFile = new File("G:/Desktop/demo.xlsx");
		if (!oldZipFile.exists()) {
			throw new RuntimeException("目标xlsx文件不存在！");
		}
		//去除密码后重新生成xml的存放目录（存在子目录xl/worksheets）
		String newFile = "temp/";
		//带密码保护的工作表集合（用于构建新的xlsx文件）
		List<ZipArchiveEntry> zipList = new ArrayList<>();
		//工作表总数
		int sheetCount = 0;
		//处理成功的工作表数
		int processCount = 0;
		//计算用时
		long begin = System.currentTimeMillis();
		//创建一个临时文件
		File tempFile = File.createTempFile(oldZipFile.getName(), null);
		if (tempFile.exists()) {
			//删除
			tempFile.delete();
		}
		//重命名为临时文件名字
		if (!oldZipFile.renameTo(tempFile)){
			//没有重命名成功
		    throw new RuntimeException("重命名失败！");
		}
		try(
				//以zip流读取xlsx文件
				ZipFile	rootZip = new ZipFile(tempFile);
				//输出流，压缩到新文件（名字与原来一致）
				ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(oldZipFile));
				) {
			//遍历
			Enumeration enu = rootZip.getEntries();
			while (enu.hasMoreElements()) {
				//使用commons-compress-1.13.jar 包中的ZipArchiveEntry代替java.util.zip包中的ZipEntry
				ZipArchiveEntry zae = (ZipArchiveEntry) enu.nextElement();
				//由于直接操作zip文件，可以通过其name属性进行判断
				if (zae.getName().startsWith("xl/worksheets/") && zae.getName().endsWith(".xml")) {
					sheetCount ++;
					try(
							//读取
							InputStream is = rootZip.getInputStream(zae);
							) {
						//构建xml文档
						Document doc = new SAXReader().read(is);
						//查找关键节点
						Element neede = doc.getRootElement().element("sheetProtection");
						if (neede != null) {
							//除去该节点
							if (doc.getRootElement().remove(neede)) {
								processCount++;
								//重新构建xml
								File f = new File(newFile+zae.getName());
								XMLWriter writer = null;
								try(
									//写入流,生成xml文件
									FileOutputStream fos = new FileOutputStream(f);
										) {
									writer = new XMLWriter(fos);
									writer.write(doc);
									writer.flush();
									zipList.add(zae);
									//添加到新的zip中(此处单个处理，亦可集中处理)
									addToZip(zos,f,"xl/worksheets/");
									//删除
									f.delete();
								} catch (IOException e) {
								}finally {
									if (writer != null) {
										writer.close();
									}
								}
							}
						}
					} catch (IOException e) {
					} catch (DocumentException e) {
					}
				}
			}
			
			//重新遍历一次,把没有的放入新的（还原xlsx文件）
			enu = rootZip.getEntries();
			while (enu.hasMoreElements()) {
				ZipArchiveEntry zae = (ZipArchiveEntry) enu.nextElement();
				if (!zipList.contains(zae)) {
					//添加到压缩文件
					zos.putNextEntry(new ZipArchiveEntry(zae.getName()));
					byte[] buf = new byte[1024];
					int len;
					try(
							InputStream is = rootZip.getInputStream(zae);
							) {
						while ((len = is.read(buf)) > 0) {
							zos.write(buf, 0, len);
						}
						zos.flush();
					} catch (IOException e) {
					}
				}
			}
			long end = System.currentTimeMillis();
			//结果
			String timestamp = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(new Date());
			System.out.println("-----【"+timestamp+"】-----");
			System.out.println("工作表总数:" + sheetCount);
			System.out.println("已处理带密码保护工作表数:" + processCount);
			System.out.println("-----【用时:"+(end-begin)+"毫秒】-----");
		} catch (IOException e) {
			System.out.println("目标文件格式有误，请确保文件为07及以上版本（.xlsx）！");
		} 
		
	}
	
	/**
	 * 把指定xml文件添加到压缩包
	 * @param zos
	 * @param addFile
	 * @param name
	 */
	private static void addToZip(ZipOutputStream zos,File addFile,String name){
		try(
				//读入文件
				InputStream is = new FileInputStream(addFile);
				) {
			byte[] buf = new byte[1024];
			int len;
			//压缩
			zos.putNextEntry(new ZipArchiveEntry(name+addFile.getName()));
			while ((len = is.read(buf)) > 0) {
				zos.write(buf, 0, len);
			}
		} catch (FileNotFoundException e) {
		} catch (IOException e) {
		}
	}
}
