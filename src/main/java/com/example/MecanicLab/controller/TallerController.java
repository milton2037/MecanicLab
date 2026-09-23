package com.example.MecanicLab.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

@Controller
@RequestMapping("/taller")
public class TallerController {

	@GetMapping
	public String inicio() {
		return "taller/inicio";
	}

	@GetMapping("/diagnostico")
	public String diagnostico() {
		return "taller/diagnostico";
	}

	@GetMapping("/cotizacion")
	public String cotizacion() {
		return "taller/cotizacion";
	}

	@GetMapping("/kanban")
	public String kanban() {
		return "taller/kanban";
	}

	@GetMapping("/asignacion")
	public String asignacion() {
		return "taller/asignacion";
	}

	@GetMapping("/inventario")
	public String inventario() {
		return "taller/inventario";
	}

	@GetMapping("/proveedores")
	public String proveedores() {
		return "taller/proveedores";
	}
}
