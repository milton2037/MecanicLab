package com.example.MecanicLab.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

@Controller
@RequestMapping("/cda")
public class CdaController {

	@GetMapping
	public String inicio() {
		return "cda/inicio";
	}

	@GetMapping("/programacion")
	public String programacion() {
		return "cda/programacion";
	}

	@GetMapping("/inspeccion")
	public String inspeccion() {
		return "cda/inspeccion";
	}

	@GetMapping("/equipos")
	public String equipos() {
		return "cda/equipos";
	}

	@GetMapping("/resultado-fur")
	public String resultadoFur() {
		return "cda/resultado-fur";
	}
}
