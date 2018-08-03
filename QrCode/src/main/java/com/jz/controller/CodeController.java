package com.jz.controller;

import java.awt.image.BufferedImage;
import java.io.IOException;
import java.io.OutputStream;

import javax.imageio.ImageIO;
import javax.servlet.http.HttpServletResponse;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import com.jz.util.CodeUtil;

import io.swagger.annotations.Api;

@Controller
@RequestMapping("/code")
@Api(value = "/code", tags = "code接口")
public class CodeController {
	
	private static final Logger logger = LoggerFactory.getLogger(CodeController.class);
	
	
	@GetMapping(value = "/code")
	public String getCode(HttpServletResponse response){
	    OutputStream os = null;
	    try {
	        // 获取图片
	        Object[] img = CodeUtil.CreateCode();
	        System.out.println("code="+img[1].toString());
	        BufferedImage image = (BufferedImage) img[0];
	        // 输出到浏览器
	        response.setContentType("image/png");
	        os = response.getOutputStream();
	        ImageIO.write(image, "png", os);
	        os.flush();
	        // 用于验证的字符串存入session
	        //this.getSession().setAttribute(Const.AUTH_CODE, img[1].toString());
	    } catch (IOException e) {
	        logger.error("验证码输出异常",e);
	    }finally {
	        if(os != null) {
	            try {
	                os.close();
	            } catch (IOException e) {
	                System.out.println("close error");
	                e.printStackTrace();
	            }
	        }
	    }
	    return null;
	}
	
	

}
