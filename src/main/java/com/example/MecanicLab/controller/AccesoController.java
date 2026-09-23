package com.example.MecanicLab.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;

@Controller
public class AccesoController {

	@GetMapping({"/", "/login"})
	public String login() {
		return "login";
	}

	@PostMapping("/login")
	public String entrarInterno() {
		return "redirect:/dashboard";
	}

	@GetMapping("/dashboard")
	public String dashboard() {
		return "dashboard";
	}
}
