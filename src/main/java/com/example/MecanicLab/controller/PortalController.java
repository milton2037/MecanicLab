package com.example.MecanicLab.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

@Controller
@RequestMapping("/portal")
public class PortalController {

	@GetMapping
	public String inicio() {
		return "portal/inicio";
	}

	@GetMapping("/citas")
	public String citas() {
		return "portal/citas";
	}

	@GetMapping("/seguimiento")
	public String seguimiento() {
		return "portal/seguimiento";
	}

	@GetMapping("/aprobaciones")
	public String aprobaciones() {
		return "portal/aprobaciones";
	}

	@GetMapping("/tienda")
	public String tienda() {
		return "portal/tienda";
	}

	@GetMapping("/pagos")
	public String pagos() {
		return "portal/pagos";
	}

	@GetMapping("/cuenta")
	public String cuenta() {
		return "portal/cuenta";
	}
}
