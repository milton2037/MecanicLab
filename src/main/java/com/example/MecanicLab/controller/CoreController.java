package com.example.MecanicLab.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

@Controller
@RequestMapping("/core")
public class CoreController {

	@GetMapping
	public String inicio() {
		return "core/inicio";
	}

	@GetMapping("/usuarios")
	public String usuarios() {
		return "core/usuarios";
	}

	@GetMapping("/clientes-vehiculos")
	public String clientesVehiculos() {
		return "core/clientes-vehiculos";
	}

	@GetMapping("/sedes-agenda")
	public String sedesAgenda() {
		return "core/sedes-agenda";
	}

	@GetMapping("/recepcion")
	public String recepcion() {
		return "core/recepcion";
	}

	@GetMapping("/seguimiento")
	public String seguimiento() {
		return "core/seguimiento";
	}

	@GetMapping("/notificaciones")
	public String notificaciones() {
		return "core/notificaciones";
	}

	@GetMapping("/pagos-facturacion")
	public String pagosFacturacion() {
		return "core/pagos-facturacion";
	}

	@GetMapping("/hoja-vida")
	public String hojaVida() {
		return "core/hoja-vida";
	}

	@GetMapping("/reportes")
	public String reportes() {
		return "core/reportes";
	}
}
