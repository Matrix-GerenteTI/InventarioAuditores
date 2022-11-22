<?php
date_default_timezone_set("America/Mexico_City");
require_once $_SERVER['DOCUMENT_ROOT']."/InventarioAuditores/Controllers/prepareExcel.php";
require_once $_SERVER['DOCUMENT_ROOT']."/InventarioAuditores/Models/inventarios.php";
// set_time_limit(0);
// ini_set('memory_limit', '20000M');

class ReporteInventarioInmediato extends PrepareExcel 
{
    public function generarReporte()
    {
        //
        // var_dump("staaap");
            $inventarios = new Inventario;
            $productos_marcados = $inventarios->obetenerInventarioSeleccionado();
            $hoy = new DateTime();
            $hoy = $hoy->format('d/m/Y');

            
        //

        $i = 9;
        $this->libro->getActiveSheet()->mergeCells("A5:E5");
        $this->libro->getActiveSheet()->setCellValue("A5", "INVENTARIO ". "ALEATORIO" ." REALIZADO EL ". $hoy);
        $this->libro->getActiveSheet()->getStyle("A5")->applyFromArray($this->labelBold);
        $this->libro->getActiveSheet()->getStyle("A5")->applyFromArray($this->centrarTexto);

        $this->libro->getActiveSheet()->setCellValue("A8", "CODIGO");
        $this->libro->getActiveSheet()->setCellValue("B8", "DESCRIPCION");
        $this->libro->getActiveSheet()->setCellValue("C8", "COSTO");
        $this->libro->getActiveSheet()->setCellValue("D8", "FAMILIA");
        $this->libro->getActiveSheet()->setCellValue("E8", "SUBFAMILIA");
        $this->libro->getActiveSheet()->setCellValue("F8", "REVISION 1");
        $this->libro->getActiveSheet()->setCellValue("G8", "REVISION 2");
        $this->libro->getActiveSheet()->setCellValue("H8", "REVISION 3");
        $this->libro->getActiveSheet()->setCellValue("I8", "SISTEMA");
        $this->libro->getActiveSheet()->setCellValue("J8", "DIFERENCIA");
        $this->libro->getActiveSheet()->setCellValue("K8", "COSTO x DIFERENCIA");
        // $this->libro->getActiveSheet()->setCellValue("L8", "HORA REALIZACION");
        $this->libro->getActivesheet()->setCellValue("L8", "REALIZADO POR:");
        // $this->libro->getActiveSheet()->mergeCells("L8:M8");
        $this->libro->getActiveSheet()->getStyle("A8:L8")->applyFromArray($this->labelBold);


        foreach ($productos_marcados as $producto) { 
        //     // extract($producto);
            $this->libro->getActiveSheet()->setCellValue("A$i",  $producto['codigo']);
            $this->libro->getActiveSheet()->setCellValue("B$i", $producto['descripcion']);
            $this->libro->getActiveSheet()->setCellValue("C$i", $inventarios->getCostoArticulo($producto['codigo']) );
            $this->libro->getActiveSheet()->getStyle("C$i")->getNumberFormat()->setFormatCode("$#,##0.00;-$#,##0.00");
            $this->libro->getActiveSheet()->setCellValue("D$i", $producto['familia']);
            $this->libro->getActiveSheet()->setCellValue("E$i", $producto['subfamilia']);
            $this->libro->getActiveSheet()->setCellValue("F$i", $producto['fisico']);
            $this->libro->getActiveSheet()->setCellValue("G$i", $producto['fisico2'] == '' ? $producto['fisico'] : $producto['fisico2']);
            $this->libro->getActiveSheet()->setCellValue("H$i", ( $producto['fisico3'] == '' ) ? "=G$i" : $producto['fisico3']);
            $this->libro->getActiveSheet()->setCellValue("I$i", $producto['stock']);
            // $this->libro->getActiveSheet()->setCellValue("L$i", $hoy);
            $this->libro->getActiveSheet()->setCellValue("L$i", $producto['auditor']);
            // $this->libro->getActiveSheet()->getStyle("L$i")->applyFromArray($this->centrarTexto);
            // $this->libro->getActiveSheet()->mergeCells("L$i:M$i");


            if ($producto['fisico2']  == 0  ) {
                if ( $producto['fisico3'] == 0) {
                    $this->libro->getActiveSheet()->setCellValue("J$i","=F$i-I$i");
                }
            }elseif( $producto['fisico3'] == 0){
                $this->libro->getActiveSheet()->setCellValue("J$i","=G$i-I$i");
            }else{
                $this->libro->getActiveSheet()->setCellValue("J$i","=H$i-I$i");
            }
            $this->libro->getActiveSheet()->setCellValue("K$i","=C$i*J$i");      
            $this->libro->getActiveSheet()->getStyle("K$i")->getNumberFormat()->setFormatCode("$#,##0.00;-$#,##0.00");
            $i++;

        }
        $this->putLogo("F1",300,150);
        $this->libro->getActiveSheet()->getStyle("A8:L".($i-1))->applyFromArray($this->bordes);
        $this->libro->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
        $this->libro->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
        $this->libro->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);
        $this->libro->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);
        $this->libro->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);
        $this->libro->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);
        $this->libro->getActiveSheet()->getColumnDimension('G')->setAutoSize(true);
        $this->libro->getActiveSheet()->getColumnDimension('H')->setAutoSize(true);
        $this->libro->getActiveSheet()->getColumnDimension('I')->setAutoSize(true);
        $this->libro->getActiveSheet()->getColumnDimension('J')->setAutoSize(true);
        $this->libro->getActiveSheet()->getColumnDimension('K')->setAutoSize(true);
        // $this->libro->getActiveSheet()->getColumnDimension('N')->setAutoSize(true); 
        $this->libro->getActiveSheet()->getColumnDimension('L')->setAutoSize(true); 

        $reporteTerminado = new \PHPExcel_Writer_Excel2007( $this->libro);
        // ob_end_clean();
        $reporteTerminado->setPreCalculateFormulas(true);
        $reporteTerminado->save('reporteInventariosAuditoresInmediato.xlsx');

        // ob_start();
        // $reporteTerminado->save('php://output');
        // $data = ob_get_contents();
        // ob_end_clean();
        
    }
}
$reporte = new ReporteInventarioInmediato;
$reporte->generarReporte();

$configCorreo = array("descripcionDestinatario" => "Reporte de Inventarios Auditores",
                    "mensaje" => "...",
                    "pathFile" => 'reporteInventariosAuditoresInmediato.xlsx',
                    "subject" => "Reporte de Inventarios Auditores",
                    //"correos" => array('gerenteti@matrix.com.mx', "raulmatrixxx@hotmail.com","gtealmacen@matrix.com.mx","gerenteventas@matrix.com.mx","dispersion@matrix.com.mx","almacenlaureles@matrix.com.mx","software2@matrix.com.mx","cavim@matrix.com.mx","admonrh@matrix.com.mx","rhmatrix2019@gmail.com"
                    // "correos" => array('gerenteti@matrix.com.mx', "raulmatrixxx@hotmail.com","gtealmacen@matrix.com.mx","gerenteventas@matrix.com.mx","almacenlaureles@matrix.com.mx","software2@matrix.com.mx","cavim@matrix.com.mx","admonrh@matrix.com.mx","rhmatrix2019@gmail.com","gerenteventasnorte@matrix.com.mx")
                    "correos" => array('mostspeed7@gmail.com')
                    );

$reporte->enviarReporte($configCorreo);