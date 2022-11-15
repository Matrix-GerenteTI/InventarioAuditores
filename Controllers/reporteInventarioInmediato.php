<?php
date_default_timezone_set("America/Mexico_City");
require_once $_SERVER['DOCUMENT_ROOT']."/InventarioAuditores/Controllers/prepareExcel.php";
require_once $_SERVER['DOCUMENT_ROOT']."/InventarioAuditores/Models/inventarios.php";

class ReporteInventarioInmediato extends PrepareExcel 
{
    public function generarReporte($productos_marcados)
    {
        $this->libro->getActiveSheet()->mergeCells("A5:D5");
                $this->libro->getActiveSheet()->setCellValue("A5", "INVENTARIO ". "ALEATORIO" ." REALIZADO EL "."14 de Noviembre de 2022");
                $this->libro->getActiveSheet()->getStyle("A5")->applyFromArray($this->labelBold);
                $this->libro->getActiveSheet()->getStyle("A5")->applyFromArray($this->centrarTexto);

        $reporteTerminado = new \PHPExcel_Writer_Excel2007( $this->libro);
        // ob_end_clean();
        $reporteTerminado->setPreCalculateFormulas(true);
        $reporteTerminado->save("reporteInventariosInmediaroAuditores.xlsx");
    }
}