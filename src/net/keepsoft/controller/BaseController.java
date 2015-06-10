package net.keepsoft.controller;


import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import net.keepsoft.server.BaseServer;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;
import org.springframework.web.servlet.ModelAndView;

@Controller
@RequestMapping("base")
public class BaseController {
	String path = "";
	@Resource
	private BaseServer bs;

	public BaseServer getBs() {
		return bs;
	}
	/**
	 * test
	 * @return
	 */
	@RequestMapping(value="/test")
	@ResponseBody
	public void findCheckListByCategoryId(String id){
		bs.test(id);
	}
	/**
	 * upload
	 * @return
	 */
	@RequestMapping(value="/upload")
	@ResponseBody
	public ModelAndView uploadFile(String fileName,HttpServletRequest request){
		
		path = request.getSession().getServletContext().getRealPath("")+"\\test\\";
		MultipartHttpServletRequest multipartRequest = (MultipartHttpServletRequest) request;     
        // 获得文件：     
        MultipartFile file = multipartRequest.getFile("file");   
        try {
			file.transferTo(new File(path+"temp.xls"));
		} catch (IllegalStateException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} 
        return new ModelAndView("redirect:/fileupload.jsp");
	}
	/**
	 * upload
	 * @throws IOException 
	 */
	@RequestMapping(value="/download")
	@ResponseBody
	public void downloadFile(String fileName,HttpServletRequest request,HttpServletResponse response) throws IOException{
		path = request.getSession().getServletContext().getRealPath("")+"\\test\\";
		String filepath=path+fileName;
		response.setContentType("multipart/form-data");  
        //2.设置文件头：最后一个参数是设置下载文件名(假如我们叫a.pdf)  
        response.setHeader("Content-Disposition", "attachment;fileName="+fileName);  
		OutputStream outputStream = response.getOutputStream();
        InputStream inputStream = new FileInputStream(filepath);
         byte[] buffer = new byte[1024];
         int i = -1;
         while ((i = inputStream.read(buffer)) != -1) {
          outputStream.write(buffer, 0, i);
         }
         outputStream.flush();
         outputStream.close();
	}
	

}
