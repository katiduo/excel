package cn.excel.controller;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.*;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;

import cn.excel.util.CommUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import cn.excel.util.ExcelUtil;




/**
 * Excel导入导出操作
 * @author songhj
 *
 */
@RestController
public class IndexViewController {

	@PostMapping("index")
	public Object course(HttpSession session){
		return ExcelUtil.process;
	}
	@RequestMapping("Initial")
	public void Initial(){
		ExcelUtil.process=null;
	}
	/**
	 * Excel数据导入
	 * @param request
	 * @param
	 * @param
	 * @return
	 */
	@RequestMapping("uploads")
	public void uploadsExcel(@RequestParam("file") MultipartFile file,HttpServletRequest request) {
		System.err.println("-------------------上传");
		if (file.isEmpty()) {
			System.err.println("文件为空");
		}
		// 获取文件名
		String fileName = file.getOriginalFilename();
		System.err.println("上传的文件名为：" + fileName);
		// 获取文件的后缀名
		String suffixName = fileName.substring(fileName.lastIndexOf("."));
		System.err.println("上传的后缀名为：" + suffixName);
		// 文件上传后的路径
		String filePath = request.getServletContext().getRealPath("//webapp//");
		// 解决中文问题，liunx下中文路径，图片显示问题
		// fileName = UUID.randomUUID() + suffixName;
		File dest = new File(filePath + fileName);
		// 检测是否存在目录
		if (!dest.getParentFile().exists()) {
			dest.getParentFile().mkdirs();
		}
		try {
			file.transferTo(dest);
			System.out.println("上传成功");
		} catch (IllegalStateException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}
	
	/*@RequestMapping("/download")
	public String downloadFile(HttpServletRequest request, HttpServletResponse response) {
		String fileName = "d.xls";
		if (fileName != null) {
			// 当前是从该工程的WEB-INF//File//下获取文件(该目录可以在下面一行代码配置)然后下载到C:\\users\\downloads即本机的默认下载的目录
			 String realPath =request.getServletContext().getRealPath("");
			String fileDirPath = new String("src/main/resources/static/");

			File fileDir = new File(fileDirPath);
			if(!fileDir.exists()){
				fileDir.mkdirs();
			}
			File targetFile = new File(fileDir.getAbsolutePath(), fileName);
			*//*String realPath = "D://test//";*//*
			File file = new File(realPath, fileName);
			if (targetFile.exists()) {
				response.setContentType("application/force-download");// 设置强制下载不打开
				response.addHeader("Content-Disposition", "attachment;fileName=" + fileName);// 设置文件名
				byte[] buffer = new byte[1024];
				FileInputStream fis = null;
				BufferedInputStream bis = null;
				try {
					fis = new FileInputStream(targetFile);
					bis = new BufferedInputStream(fis);
					OutputStream os = response.getOutputStream();
					int i = bis.read(buffer);
					while (i != -1) {
						os.write(buffer, 0, i);
						i = bis.read(buffer);
					}
					System.out.println("success");
				} catch (Exception e) {
					e.printStackTrace();
				} finally {
					if (bis != null) {
						try {
							bis.close();
						} catch (IOException e) {
							e.printStackTrace();
						}
					}
					if (fis != null) {
						try {
							fis.close();
						} catch (IOException e) {
							e.printStackTrace();
						}
					}
				}
			}
		}

		return null;
	}
	*/
	
	@RequestMapping(value="/importExcel")
	public Map<String, Object> importExcel(HttpServletRequest request,HttpServletResponse response, String filePro){
		Map<String, Object> map = new HashMap<>();
		String [] keys = {"phone","shopName"};
		System.err.println(""+filePro);
		try {

			List<Map<String,String>> listData = ExcelUtil.getExcelData(request, "",keys);
			if(listData.size() == 0){
				map.put("status",-1);
				map.put("message","上传失败，上传数据必须大于一条");
				return map;
			}
			for (Map<String, String> dataMap : listData) {
				System.out.println(keys[0] + ":" + dataMap.get(keys[0]));
				System.out.println(keys[1] + ":" + dataMap.get(keys[1]));
			}
			map.put("listData", listData);
			map.put("code", 1);
			map.put("message", "导入成功");
			ExcelUtil.process=null;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return map;
	}

	/**
	 * 数据导出Excel 模板
	 * @param request
	 * @param response
	 * @throws IOException
	 */
	@RequestMapping(value = "/excelExport", method = { RequestMethod.GET, RequestMethod.POST })
	public String excelExport(HttpServletRequest request, HttpServletResponse response) throws IOException {
		/*String fileName = CommUtil.formatTime("yyyyMMddHHmmss", new Date()) +".xlsx";*/

		String fileName="model.xlsx";

		String [] columnNames = {"手机号","店铺名"};
		String [] keys = {"phone","shopName"};

		List<Map<String,Object>> listMap = new ArrayList<>();
		Map<String,Object> map = new HashMap<>();
		for (int i=0;i<60000;i++){
			listMap.add(map);
		}
		try {
			//保存路径
			String fileDirPath = new String("src/main/resources/static/model");
			Workbook wb ;
			File fileDir = new File(fileDirPath);
			File targetFile = new File(fileDir.getAbsolutePath(), fileName);
			if(!fileDir.exists()){
				fileDir.mkdirs();
				//创建Workbook
				wb = ExcelUtil.createWorkBook(listMap, keys, columnNames);
			}else{
				FileInputStream tps = new FileInputStream(targetFile);
				wb = new XSSFWorkbook(tps);
			}

			excelDownldBody(wb,targetFile,response,fileName);
		/*	//返回结果
			data.put("code", 1);
			String downloadUrl = request.getScheme() + "://"+request.getServerName() + ":" +
					request.getServerPort() + "/"+ fileName;
			data.put("download", downloadUrl);
			data.put("message", "文件流输出成功");

			System.out.println("\n数据导出成功，下载路劲：" + downloadUrl);*/
		} catch (Exception e) {
			System.err.println(e.getMessage());
			/*data.put("code", -1);
			data.put("message", "下载出错");
			return data;*/
		}
		return null;
	}
	/**
	* 根据excel模板导出数据
	* */
	@RequestMapping("/modelOut")
	public String copy(HttpServletResponse response,@RequestParam("name") String name){
		try {
			//创建Workbook
			List<Map<String,String>> list=new ArrayList<>();
			Map<String,String> map=new HashMap<>();

			map.put("phone","17812345678");
			map.put("shop","京东");
			list.add(map);
			map.put("phone","17812345678");
			map.put("shop","京东");
			list.add(map);
			map.put("phone","17812345678");
			map.put("shop","京东");
			list.add(map);

			//保存路径
			String fileDirPath = new String("D:/ChromeCoreDownloads/");

			File fileDir = new File(fileDirPath);
			if(!fileDir.exists()){
				fileDir.mkdirs();
			}
			String fileName=name+".xlsx";
			File targetFile = new File(fileDir.getAbsolutePath(),fileName );
/*HSSF
			String savePath = request.getServletContext().getRealPath("/") + File.separator + fileName;
*/
			FileInputStream tps = new FileInputStream(targetFile);
			final Workbook wb = new XSSFWorkbook(tps);
			Sheet sheet=wb.getSheetAt(0);

			for(int i=0;i<list.size();i++){
				int j=0;
				Row row=sheet.getRow(i+1);
				for(Map.Entry<String, String> m:list.get(i).entrySet()){
						Cell cell=row.getCell(j);
						cell.setCellValue(m.getValue());
						j++;
				}
			}
			excelDownldBody(wb,targetFile,response,fileName);
		} catch (IOException e) {
			System.err.println(e.getMessage());
		}
		return null;
	}
	/**
	 * 下载的主体部分
	 * */
	private void excelDownldBody(Workbook wb,File targetFile,HttpServletResponse response,String fileName){
		try {
			// 创建文件流
			OutputStream stream = new FileOutputStream(targetFile);
			// 写入数据
			wb.write(stream);
			// 关闭文件流
			stream.close();
			if (targetFile.exists()) {
				response.setContentType("application/force-download");// 设置强制下载不打开
				response.addHeader("Content-Disposition", "attachment;fileName=" + fileName);// 设置文件名
				byte[] buffer = new byte[1024];
				FileInputStream fis = null;
				BufferedInputStream bis = null;
				try {
					fis = new FileInputStream(targetFile);
					bis = new BufferedInputStream(fis);
					OutputStream os = response.getOutputStream();
					int i = bis.read(buffer);
					while (i != -1) {
						os.write(buffer, 0, i);
						i = bis.read(buffer);
					}
					System.out.println("success");
				} catch (Exception e) {
					e.printStackTrace();
				} finally {
					if (bis != null) {
						try {
							bis.close();
						} catch (IOException e) {
							e.printStackTrace();
						}
					}
					if (fis != null) {
						try {
							fis.close();
						} catch (IOException e) {
							e.printStackTrace();
						}
					}
				}
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
