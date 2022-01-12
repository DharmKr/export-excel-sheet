package com.ExportExcelSheet.service;

import java.io.IOException;
import java.util.List;

import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.MessagingException;
import javax.mail.internet.MimeMessage;
import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.stereotype.Service;

import com.ExportExcelSheet.entity.UserEntity;
import com.ExportExcelSheet.repository.UserRepository;

@Service
public class UserService {
	
	@Autowired
	private UserRepository userRepository;
	
	@Autowired
	private JavaMailSender javaMailSender;

	public List<UserEntity> getUserList() throws IOException, MessagingException{
		List<UserEntity> userList = userRepository.findAll();
		exportToExcel(userList);
		return userList;
	}
	
	private void exportToExcel(List<UserEntity> userLists) throws IOException, MessagingException {
        UserExcelExporterService excelExporter = new UserExcelExporterService(userLists);
        byte[] excelFileAsBytes = excelExporter.export();
        ByteArrayResource resource = new ByteArrayResource(excelFileAsBytes);
		sendMail(resource);
	}
	
	public void sendMail(ByteArrayResource resource) throws MessagingException, IOException {
		MimeMessage message = javaMailSender.createMimeMessage();
		MimeMessageHelper helper = new MimeMessageHelper(message, true);
		helper.setFrom("dharmender.mytoshika@gmail.com");
		helper.setTo("dharmenderkumarabc123@gmail.com");
		helper.setSubject("Test Mail for sending excel sheet");
		helper.setText("UserDetails excel sheet-");
		helper.addAttachment("UserDetails.xlsx", resource);
		javaMailSender.send(message);
	}
}
