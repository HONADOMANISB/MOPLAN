<?php

namespace App\Http\Controllers\Ceplan;


use App\Http\Controllers\Respuesta\JSONResponseController;
use App\Models\Ceplan\CeplanModel;
use Carbon\Carbon;
use Illuminate\Http\Request;
use Mpdf\Mpdf;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use Illuminate\Support\Facades\Auth;
class CeplanController extends JSONResponseController

{
    public function __construct() {
        $this->middleware('auth:sanctum');
    }
    public function uploadExcel(Request $request)
    {

        $file = $request->file('archivo');
        $excel = IOFactory::load($file);

        $sheet = $excel->getActiveSheet();
        $highestRow = $sheet->getHighestDataRow();
        $date = Carbon::now();  
        $tipo = $request->post('tipo');
        $data=[];
        $vnp=[];
        $resultado = [];
        $año=$sheet->getCell('A2')->getCalculatedValue();
            for ($row = 2; $row <= $highestRow; $row++) {

                $data = [
                    'YEAR' => $sheet->getCell('A' . $row)->getCalculatedValue(),
                    'ETAPA' => $sheet->getCell('B' . $row)->getCalculatedValue(),
                    'UE_ID' => $sheet->getCell('C' . $row)->getCalculatedValue(),
                    'UE' => $sheet->getCell('D' . $row)->getCalculatedValue(),
                    'CC_RESPONSABLE_ID' => $sheet->getCell('E' . $row)->getCalculatedValue(),
                    'DEPARTAMENTO'=> $sheet->getCell('F' . $row)->getCalculatedValue(),
                    'CENTRO_COSTOS_ID' => $sheet->getCell('G' . $row)->getCalculatedValue(),
                    'CENTRO_COSTOS' => $sheet->getCell('H' . $row)->getCalculatedValue(),
                    'SERVICIO'=> $sheet->getCell('I' . $row)->getCalculatedValue(),
                    'USUARIO' => $sheet->getCell('J' . $row)->getCalculatedValue(),
                    'DATOS_USUARIO' => $sheet->getCell('K' . $row)->getCalculatedValue(),
                    'OEI' => $sheet->getCell('L' . $row)->getCalculatedValue(),
                    'OBJETIVO_ESTRATEGICO' => $sheet->getCell('M' . $row)->getCalculatedValue(),
                    'AEI' => $sheet->getCell('N' . $row)->getCalculatedValue(),
                    'ACCION_ESTRATEGICA' => $sheet->getCell('O' . $row)->getCalculatedValue(),
                    'CATEGORIA_ID' => $sheet->getCell('P' . $row)->getCalculatedValue(),
                    'CATEGORIA' => $sheet->getCell('Q' . $row)->getCalculatedValue(),
                    'PRODUCTO_ID' => $sheet->getCell('R' . $row)->getCalculatedValue(),
                    'PRODUCTO' => $sheet->getCell('S' . $row)->getCalculatedValue(),
                    'FUNCION_ID' => $sheet->getCell('T' . $row)->getCalculatedValue(),
                    'FUNCION' => $sheet->getCell('U' . $row)->getCalculatedValue(),
                    'DIVISION_FUNCIONAL_ID' => $sheet->getCell('V' . $row)->getCalculatedValue(),
                    'DIVISION_FUNCIONAL' => $sheet->getCell('W' . $row)->getCalculatedValue(),
                    'GRUPO_FUNCIONAL_ID' => $sheet->getCell('X' . $row)->getCalculatedValue(),
                    'GRUPO_FUNCIONAL' => $sheet->getCell('Y' . $row)->getCalculatedValue(),
                    'ACTIVIDAD_PRESUPUESTAL_ID' => $sheet->getCell('Z' . $row)->getCalculatedValue(),
                    'ACTIVIDAD_PRESUPUESTAL' => $sheet->getCell('AA' . $row)->getCalculatedValue(),
                    'NRO_REGISTRO_POI' => $sheet->getCell('AB' . $row)->getCalculatedValue(),
                    'ACTIVIDAD_OPERATIVA_ID' => $sheet->getCell('AC' . $row)->getCalculatedValue(),
                    'CODIGO_PPR' => $sheet->getCell('AD' . $row)->getCalculatedValue(),
                    'ACTIVIDAD_OPERATIVA' => $sheet->getCell('AE' . $row)->getCalculatedValue(),
                    'UNIDAD_MEDIDA' => $sheet->getCell('AF' . $row)->getCalculatedValue(),
                    'TRAZADORA_TAREA' => $sheet->getCell('AG' . $row)->getCalculatedValue(),
                    'ACUMULADO' => $sheet->getCell('AH' . $row)->getCalculatedValue(),
                    'DEFINICION_OPERACIONAL' => $sheet->getCell('BZ' . $row)->getCalculatedValue(),   
                    'ESTADO_ACTIVIDAD_OPERATIVA' => $sheet->getCell('CA' . $row)->getCalculatedValue(),
                    'TIPO_ACTIVIDAD' => $sheet->getCell('CB' . $row)->getCalculatedValue(),  
                    'TIPO_REGISTRO' => $sheet->getCell('CC' . $row)->getCalculatedValue(),  
                    'ACTIVIDAD_CT' => $sheet->getCell('CD' . $row)->getCalculatedValue(),  
                    'CENTRO_COSTO_HSB' => $sheet->getCell('CE' . $row)->getCalculatedValue(),  
                    'FECHA_EXPORTA'=>$date,  
                    'SG_EST_REGISTRO'=>'A', 
                    'TIPO'=>$tipo                       
                ];
                $vnp = [
                    [
                        'MES' => '1',
                        'PROGRAMADO' =>  $sheet->getCell('AI' . $row)->getCalculatedValue(),
                        'EJECUTADO'  =>  $sheet->getCell('AV' . $row)->getCalculatedValue(),
                        'DETALLE_MOTIVO'     =>  $sheet->getCell('BI' . $row)->getCalculatedValue(),
                        'MOTIVO'     =>  $sheet->getCell('BJ' . $row)->getCalculatedValue(),
                    ],           
                    [
                        'MES' => '2',
                        'PROGRAMADO' =>  $sheet->getCell('AJ' . $row)->getCalculatedValue(),
                        'EJECUTADO'  =>  $sheet->getCell('AW' . $row)->getCalculatedValue(),
                        'DETALLE_MOTIVO'     =>  $sheet->getCell('BK' . $row)->getCalculatedValue(),
                        'MOTIVO'     =>  $sheet->getCell('BL' . $row)->getCalculatedValue(),
                    ],
                    [
                        'MES' => '3',
                        'PROGRAMADO'=> $sheet->getCell('AK' . $row)->getCalculatedValue(),
                        'EJECUTADO' =>  $sheet->getCell('AX' . $row)->getCalculatedValue(),
                        'DETALLE_MOTIVO'    =>  $sheet->getCell('BM' . $row)->getCalculatedValue(),
                        'MOTIVO'     =>  $sheet->getCell('BN' . $row)->getCalculatedValue(),
                    ],
                    [
                        'MES' => '4',
                        'PROGRAMADO' => $sheet->getCell('AL' . $row)->getCalculatedValue(),
                        'EJECUTADO'  =>  $sheet->getCell('AY' . $row)->getCalculatedValue(),
                        'DETALLE_MOTIVO'     =>  $sheet->getCell('BO' . $row)->getCalculatedValue(),
                        'MOTIVO'     =>  $sheet->getCell('BP' . $row)->getCalculatedValue(),
                    ],
                    [
                        'MES' => '5',
                        'PROGRAMADO' => $sheet->getCell('AM' . $row)->getCalculatedValue(),
                        'EJECUTADO' =>  $sheet->getCell('AZ' . $row)->getCalculatedValue(),
                        'DETALLE_MOTIVO'    =>  $sheet->getCell('BQ' . $row)->getCalculatedValue(),
                        'MOTIVO'     =>  $sheet->getCell('BR' . $row)->getCalculatedValue(),
                    ],
                    [
                        'MES' => '6',
                        'PROGRAMADO' => $sheet->getCell('AN' . $row)->getCalculatedValue(),
                        'EJECUTADO' =>  $sheet->getCell('BA' . $row)->getCalculatedValue(),
                       
                    ],
                    [
                        'MES' => '7',
                        'PROGRAMADO' => $sheet->getCell('AO' . $row)->getCalculatedValue(),
                        'EJECUTADO' =>  $sheet->getCell('BB' . $row)->getCalculatedValue(),  
                        
                    ],
                    [
                        'MES' => '8',
                        'PROGRAMADO' => $sheet->getCell('AP' . $row)->getCalculatedValue(),
                        'EJECUTADO' =>  $sheet->getCell('BC' . $row)->getCalculatedValue(),
                        
                    ],
                    [
                        'MES' => '9',
                        'PROGRAMADO' => $sheet->getCell('AQ' . $row)->getCalculatedValue(),
                        'EJECUTADO' =>  $sheet->getCell('BD' . $row)->getCalculatedValue(),
                        
                    ],
                    [
                        'MES' => '10',
                        'PROGRAMADO' => $sheet->getCell('AR' . $row)->getCalculatedValue(),
                        'EJECUTADO' =>  $sheet->getCell('BE' . $row)->getCalculatedValue(),
                        
                    ],
                    [
                        'MES' => '11',
                        'PROGRAMADO' => $sheet->getCell('AS' . $row)->getCalculatedValue(),
                        'EJECUTADO' =>  $sheet->getCell('BF' . $row)->getCalculatedValue(), 
                      
                    ],
                    [
                        'MES' => '12',
                        'PROGRAMADO' => $sheet->getCell('AT' . $row)->getCalculatedValue(),
                        'EJECUTADO' =>  $sheet->getCell('BG' . $row)->getCalculatedValue(),
                       
                    ],
                    [
                        'MES' => '13',
                        'PROGRAMADO' => $sheet->getCell('AU' . $row)->getCalculatedValue(),
                        'EJECUTADO' =>  $sheet->getCell('BH' . $row)->getCalculatedValue(),
                    ],
                ];
                
                foreach($vnp as $vnpU){
                    $resultado[] = array_merge($vnpU, $data);
                }
            }
        
            $seguro=new CeplanModel();
            $response=$seguro->insertarExcel($resultado,$año,$tipo);
            return $this->sendResponse(200, true, $response, 1);

    }
    public function listarActividades(Request $request){
         $user = $request->user();
         $perfil = $user->id_perfil;
         $servicio=$request->post('servicio');
         $year=$request->post('cb_year');
         $periodo=$request->post('cb_mes');
         $actividad=new CeplanModel();
         $resultado=$actividad->listarActividades($servicio,$perfil,$year,$periodo);
         return $this->sendResponse(200, true,'',$resultado);
    }
    public function listarDepartamentos(Request $request){
        $actividad=new CeplanModel();
        $resultado=$actividad->listarDepartamentos();
        return $this->sendResponse(200, true,'',$resultado);
   }
   public function listarMotivos(Request $request){
    $actividad=new CeplanModel();
    $resultado=$actividad->listarMotivos();
    return $this->sendResponse(200, true,'',$resultado);
}
      public function listarServicios(Request $request){
        $departamento=$request->post('departamento');
        $actividad=new CeplanModel();
        $resultado=$actividad->listarServicios($departamento);
        return $this->sendResponse(200, true,'',$resultado);
   }
    public function listarInformacion(Request $request){
        $act=$request->post('actividad');
        $year=$request->post('cb_year');
        $actividad=new CeplanModel();
        $resultado=[];
        $resultado['PI']=$actividad->listarInformacion($act,$year);
        $resultado['PR']=$actividad->listarInformacionPR($act,$year);
        $resultado['SR']=$actividad->listarInformacionSR($act,$year);

        return $this->sendResponse(200, true,'',$resultado);
   }
   public function listarEncabezado(Request $request){
    $act=$request->post('actividad');
    $year=$request->post('cb_year');
    $actividad=new CeplanModel();
    $resultado=$actividad->listarEncabezado($act,$year);
    return $this->sendResponse(200, true,'',$resultado);
}
public function generarReporteDetallePOI(Request $request){
    $mes=$request->get('mes');
    $year=$request->get('year');
    $ppr=$request->get('ppr');
    $tipo=$request->get('tipo');
    $report=new CeplanModel();
    $response=$report->generarReporteDetallePOI($mes,$year,$ppr,$tipo);
    $html='';      
    $html.='<h5 style="text-align:center;">REPORTE DETALLE POI </h5>';
    $html.='
    <table>
      <thead>
           <tr>
            <td>#</td>  
            <td>N° EPISODIO</td> 
            <td>N° HISTORIA</td>
            <td>APELLIDOS Y NOMBRES</td>
            <td>FECHA ATENCIÓN</td>
            <td>CODIGO MÈDICO</td>
            <td> MÈDICO</td>
           </tr>
      </thead>
    <tbody> ';
    $count=1;
    foreach ($response as $data) {
        $html.='
        <tr>
         <td>'.$count++.'</td>
         <td>'.$data->NUM_EPISODIO.'</td>
         <td>'.$data->HISTORIA_CLINICA.'</td>
         <td>'.$data->NOMBRE_COMPLETO.'</td>
         <td>'.$data->FECHA_ATENCION.'</td>
         <td>'.$data->COD_MEDICO.'</td>
         <td>'.$data->MEDICO.'</td>
        </tr>    
      ';
    }
    $html.='
    </tbody>
    </table>';

    $mpdf = new Mpdf();
    $css = file_get_contents(resource_path('css\\reporteDetallePOI.css')); // css
    $header='
    <div>
    <table class="encabezado">
    <tr>
       <td colspan="3" class="pega">
         <img width = "100" src = "'.resource_path().'/img/img_personal.png" class="pegantina" alt="imagen no ecnontrada">
       </td>
       <td class="logo"><img width = "60" src = "'.resource_path().'/img/logo-hsb.jpg" class="log" alt="imagen no ecnontrada"></td>  </tr>            
    </tr> 
   </table></div>';
    $mpdf->SetMargins(15, 60,30);
    $mpdf->SetHTMLHeader($header);
    $mpdf->WriteHTML($css,1);
    $mpdf->WriteHTML($html,2);
    $mpdf->Output(); 
}
public function generarReporteDetallePOIExcel(Request $request){
    $mes=$request->get('mes');
    $year=$request->get('year');
    $ppr=$request->get('ppr');
    $tipo=$request->get('tipo');
    $report=new CeplanModel();
    $response=$report->generarReporteDetallePOI($mes,$year,$ppr,$tipo);
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $encabezado = [
        "Nª",
        "Nª EPISODIO",
        "Nª HIS",
        "Nª HISTORIA",
        "Nª PACIENTE",
        "FECHA ATENCIÒN",
        "CODIGO MEDICO",
        "MEDICO"

    ];

    $sheet->fromArray($encabezado, null, 'A1');
    $row = 2;
    $C=1;
    foreach ($response as $valor) {
        $sheet->setCellValue('A' . $row, $C++);   
        $sheet->setCellValue('B' . $row, $valor->NUM_EPISODIO);
        $sheet->setCellValue('C' . $row, $valor->NUM_HIS);
        $sheet->setCellValue('D' . $row, $valor->HISTORIA_CLINICA);
        $sheet->setCellValue('E' . $row, $valor->NOMBRE_COMPLETO);
        $sheet->setCellValue('F' . $row, $valor->FECHA_ATENCION);
        $sheet->setCellValue('G' . $row, $valor->COD_MEDICO);
        $sheet->setCellValue('H' . $row, $valor->MEDICO);       
        $row++;
    }
    for ($col = 'A'; $col <= 'H'; $col++) { $sheet->getColumnDimension($col)->setAutoSize(true);}
    $styleArray = [
        'borders' => [
            'allBorders' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['argb' => Color::COLOR_BLACK],
            ],
        ],
        'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_LEFT,
                'wrapText' => false,
            ]
    ];

    $rowFInal = $row - 1;
    $highestColumn = $sheet->getHighestColumn();
    $cellRange = 'A1:' . $highestColumn . $rowFInal;
    $sheet->getStyle($cellRange)->applyFromArray($styleArray);
    $writer = new Xlsx($spreadsheet);
    $fileName = 'reporte.xlsx';
    $writer->save($fileName);

    return response()->download($fileName)->deleteFileAfterSend(true);

}

public function invalidarPoi(Request $request){
    [$usuario, $perfil, $equipo] = $this->getHost($request);
    $ejecutado=$request->post('ej');
    $id=$request->post('mes');
    $motivo=$request->post('motivo');
    $actividad=$request->post('actividad');
    $tipo=$request->post('tipo');
    $activity=new CeplanModel();
    $response=$activity->invalidarPoi($ejecutado,$id,$motivo,$actividad,$tipo,$usuario,$equipo,$perfil);
    return $this->sendResponse(200, true,'',$response);
  
}
public function cerrarActividades(Request $request){
    [$usuario, $perfil, $equipo] = $this->getHost($request);
    $id=$request->post('mes');
    $year=$request->post('year');
    $activity=new CeplanModel();
    $response=$activity->cerrarActividades($id,$year,$usuario,$equipo,$perfil);
    return $this->sendResponse(200, true,'',$response);
  
}
   public function guardarPoi(Request $request){
    [$usuario, $perfil, $equipo] = $this->getHost($request);
    $id=$request->post('mes');
    $año=$request->post('year');
    $ejecutado=$request->post('ejecutado');
    $tipo_registro=$request->post('tipoEstado');
    $detalle_motivo=$request->post('detalleMotivo');
    $motivo=$request->post('motivo');
    $actividad=$request->post('actividad');
    $tipo=$request->post('tipo');
    $activity=new CeplanModel();

    if(!$motivo) $motivo="";
    if(!$tipo)$tipo=" ";
    $response=$activity->guardarPoi($id,$año,$ejecutado,$motivo,$actividad,$tipo,$tipo_registro,$detalle_motivo,$usuario,$equipo,$perfil);
    return $this->sendResponse(200, true,'',$response);
  
}

public function registrarLogros(Request $request){
    [$usuario, $perfil, $equipo] = $this->getHost($request);
    $user = Auth::user();
    $servicio = $user->servicio;
    $periodo=$request->post('periodo');
    $año=$request->post('año');
    $logro=$request->post('logro');
    $dificultad=$request->post('dificultad');
    $accion_mejora=$request->post('accion_mejora');
    $accion_correctiva=$request->post('accion_correctiva');
    $actividad=$request->post('actividad');
    $activity=new CeplanModel();

   
    $response=$activity->registrarLogros($periodo,$año,$logro,$dificultad,$accion_mejora,$accion_correctiva,$actividad,$servicio,$usuario,$equipo,$perfil);
    return $this->sendResponse(200, true,'',$response);
  
}
public function listarActividadesLogros(Request $request){
    $user = $request->user();
    $perfil = $user->id_perfil;
    $servicio=$request->post('servicio');
    $year=$request->post('cb_year');
    $actividad=new CeplanModel();
    $resultado=$actividad->listarActividadesLogros($servicio,$perfil,$year);
    return $this->sendResponse(200, true,'',$resultado);
}
public function listarLogros(Request $request){
    $trimestre=$request->post('cb_trimestre');
    $year=$request->post('cb_year');
    $actividad=$request->post('actividad');
    $con=new CeplanModel();
    $resultado=$con->listarLogros($trimestre,$year,$actividad);
    return $this->sendResponse(200, true,'',$resultado);
}
public function listarEventos(Request $request){
    $con=new CeplanModel();
    $resultado=$con->listarEventos();
    return $this->sendResponse(200, true,'',$resultado);
}
}
