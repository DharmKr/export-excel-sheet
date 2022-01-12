package com.ExportExcelSheet.controller;

import java.io.IOException;

import javax.mail.MessagingException;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.ExportExcelSheet.service.UserService;

@RestController
@RequestMapping("/api")
public class UserController {
	
	@Autowired
	private UserService userService;

	@GetMapping("/getUserList")
	public String getUserList() throws IOException, MessagingException{
		userService.getUserList();
		return "excel sheet created";
	}
}
