<?php
defined('BASEPATH') OR exit('No direct script access allowed');
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class Stock_balance_zone extends PS_Controller
{
  public $menu_code = 'RICSBZ';
	public $menu_group_code = 'RE';
  public $menu_sub_group_code = 'REINVT';
	public $title = 'รายงานสินค้าคงเหลือ แยกตามโซน (SAP)';
  
  public function __construct()
  {
    parent::__construct();
    $this->home = base_url().'report/inventory/stock_balance_zone';
    $this->load->model('report/inventory/stock_balance_report_model');
    $this->load->model('masters/products_model');
  }

  public function index()
  {
    $this->load->model('masters/warehouse_model');
    $whList = $this->warehouse_model->get_all_warehouse();
    $ds['whList'] = $whList;
    $this->load->view('report/inventory/report_stock_balance_zone', $ds);
  }


  public function get_report()
  {
    ini_set('memory_limit','512M'); // This also needs to be increased in some cases. Can be changed to a higher value as per need)
    ini_set('sqlsrv.ClientBufferMaxKBSize','524288'); // Setting to 512M
    ini_set('pdo_sqlsrv.client_buffer_max_kb_size','524288'); // Setting to 512M - for pdo_sqlsrv
    
    $sc = TRUE;
    $allProduct = $this->input->get('allProduct');
    $pdFrom = $this->input->get('pdFrom');
    $pdTo = $this->input->get('pdTo');

    $allWhouse = $this->input->get('allWhouse');
    $warehouse = $this->input->get('warehouse');

    $allZone = $this->input->get('allZone');
    $zoneCode = $this->input->get('zoneCode');
    $zoneName = $this->input->get('zoneName');

    $wh_list = '';
    if(!empty($warehouse) && empty($zoneCode))
    {
      $i = 1;
      foreach($warehouse as $wh)
      {
        $wh_list .= $i === 1 ? $wh : ', '.$wh;
        $i++;
      }
    }

    //---  Report title
    $ds['reportDate'] = thai_date(date('Y-m-d'),FALSE, '/');
    $ds['whList']   = $allWhouse == 1 ? 'ทั้งหมด' : $wh_list;
    $ds['zoneList'] = $allZone == 1 ? 'ทั้งหมด' : $zoneCode." - ".$zoneName;
    $ds['productList']   = $allProduct == 1 ? 'ทั้งหมด' : '('.$pdFrom.') - ('.$pdTo.')';

    $bs = array();

    $result = $this->stock_balance_report_model->get_stock_balance_zone($allProduct, $pdFrom, $pdTo, $allWhouse, $warehouse, $allZone, $zoneCode);

    if(!empty($result))
    {
      if(count($result) > 2000)
      {
        $sc = FALSE;
        $this->error = "ข้อมูลมีปริมาณมากเกินกว่าจะแสดงผลได้ กรุณาส่งออกข้อมูลแทนการแสดงผลหน้าจอ";
      }
      else
      {
        $no = 1;
        $totalQty = 0;
        $totalAmount = 0;

        foreach($result as $rs)
        {
          $amount = $rs->qty * $rs->price;

          $arr = array(
            'no' => number($no),
            'warehouse' => $rs->warehouse_code,
            'zone' => $rs->zone_name,
            'pdCode' => $rs->product_code,
            'pdName' => $rs->product_name,
            'price' => number($rs->price, 2),
            'qty' => number($rs->qty),
            'amount' => number($amount, 2)
          );

          array_push($bs, $arr);
          $totalQty += $rs->qty;
          $totalAmount += $amount;

          $no++;
        }

        $arr = array( 'totalQty' => number($totalQty), 'totalAmount' => number($totalAmount, 2));
        array_push($bs, $arr);
      }
    }
    else
    {
      $arr = array('nodata' => 'nodata');
      array_push($bs, $arr);
    }

    $ds['bs'] = $bs;

    echo $sc === TRUE ? json_encode($ds) : $this->error;
  }





  public function do_export()
  {
    ini_set('memory_limit','512M'); // This also needs to be increased in some cases. Can be changed to a higher value as per need)
    ini_set('sqlsrv.ClientBufferMaxKBSize','524288'); // Setting to 512M
    ini_set('pdo_sqlsrv.client_buffer_max_kb_size','524288'); // Setting to 512M - for pdo_sqlsrv
   
    $allProduct = $this->input->post('allProduct');
    $pdFrom = $this->input->post('pdFrom');
    $pdTo = $this->input->post('pdTo');
    $allWhouse = $this->input->post('allWhouse');
    $warehouse = $this->input->post('warehouse');
    $allZone = $this->input->post('allZone');
    $zoneCode = $this->input->post('zoneCode');
    $zoneName = $this->input->post('zoneName');
    $token = $this->input->post('token');

    $wh_list = '';

    if(!empty($warehouse) && empty($zoneCode))
    {
      $i = 1;
      foreach($warehouse as $wh)
      {
        $wh_list .= $i === 1 ? $wh : ', '.$wh;
        $i++;
      }
    }

    //---  Report title
    $report_title = "รายงานสินค้าคงเหลือแยกตามโซน ณ วันที่ ".date('d/m/Y');
    $whList = $allWhouse == 1 ? 'ทั้งหมด' : $wh_list;
    $zoneList = $allZone == 1 ? 'ทั้งหมด' : $zoneCode." - ".$zoneName;
    $productList  = $allProduct == 1 ? 'ทั้งหมด' : '('.$pdFrom.') - ('.$pdTo.')';
   
    $result = $this->stock_balance_report_model->get_stock_balance_zone($allProduct, $pdFrom, $pdTo, $allWhouse, $warehouse, $allZone, $zoneCode);

   $excel = new Spreadsheet();
   $sheet = $excel->getActiveSheet();
   $sheet->setTitle('Stock Balance Report');

    //--- set report title header
    $sheet->setCellValue('A1', $report_title);
    $sheet->mergeCells('A1:G1');
    $sheet->setCellValue('A2', 'คลัง');
    $sheet->setCellValue('B2', $whList);
    $sheet->mergeCells('B2:G2');
    $sheet->setCellValue('A3', 'โซน');
    $sheet->setCellValue('B3', $zoneList);
    $sheet->mergeCells('B3:G3');
    $sheet->setCellValue('A4', 'สินค้า');
    $sheet->setCellValue('B4', $productList);
    $sheet->mergeCells('B4:G4');

    //--- set Table header
    $sheet->setCellValue('A5', 'ลำดับ');
    $sheet->setCellValue('B5', 'รหัสคลัง');
    $sheet->setCellValue('C5', 'รหัสโซน');
    $sheet->setCellValue('D5', 'ชื่อโซน');
    $sheet->setCellValue('E5', 'รหัสสินค้า');
    $sheet->setCellValue('F5', 'ชื่อสินค้า');
    $sheet->setCellValue('G5', 'ราคาขาย');
    $sheet->setCellValue('H5', 'จำนวน');
    $sheet->setCellValue('I5', 'มูลค่า');

    $row = 6;
    if(!empty($result))
    {
      $no = 1;
      $totalQty = 0;
      foreach($result as $rs)
      {
        $sheet->setCellValue('A'.$row, $no);
        $sheet->setCellValue('B'.$row, $rs->warehouse_code);
        $sheet->setCellValue('C'.$row, $rs->zone_code);
        $sheet->setCellValue('D'.$row, $rs->zone_name);
        $sheet->setCellValue('E'.$row, $rs->product_code);
        $sheet->setCellValue('F'.$row, $rs->product_name);
        $sheet->setCellValue('G'.$row, $rs->price);
        $sheet->setCellValue('H'.$row, $rs->qty);
        $sheet->setCellValue('I'.$row, "=G{$row}*H{$row}");
        $no++;
        $row++;
      }

      $ro = $row - 1;
      $sheet->setCellValue('A'.$row, 'รวม');
      $sheet->mergeCells('A'.$row.':G'.$row);
      $sheet->setCellValue('H'.$row, "=SUM(H6:H{$ro})");
      $sheet->setCellValue('I'.$row, "=SUM(I6:I{$ro})");
      $sheet->getStyle('A'.$row)->getAlignment()->setHorizontal('right');
    }

    setToken($token);
    $file_name = "Report Stock Zone.xlsx";
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); /// form excel 2007 XLSX
    header('Content-Disposition: attachment;filename="'.$file_name.'"');
    $writer = new Xlsx($excel);
    $writer->save('php://output');
  }


  public function export_to_check()
  {
    ini_set('memory_limit','512M'); // This also needs to be increased in some cases. Can be changed to a higher value as per need)
    ini_set('sqlsrv.ClientBufferMaxKBSize','524288'); // Setting to 512M
    ini_set('pdo_sqlsrv.client_buffer_max_kb_size','524288'); // Setting to 512M - for pdo_sqlsrv

    $allProduct = 1;
    $pdFrom = NULL;
    $pdTo = NULL;

    $allWhouse = 1;
    $warehouse = NULL;

    $allZone = 0;
    $zoneCode = $this->input->post('zone_code');
    $token = $this->input->post('token');

    $result = $this->stock_balance_report_model->get_stock_balance_zone($allProduct, $pdFrom, $pdTo, $allWhouse, $warehouse, $allZone, $zoneCode);

    $excel = new Spreadsheet();
    $sheet = $excel->getActiveSheet();
    $sheet->setTitle('Report Stock Zone');

    //--- set Table header
    $sheet->setCellValue('A1', 'barcode');
    $sheet->setCellValue('B1', 'item_code');
    $sheet->setCellValue('C1', 'qty');

    $row = 2;
    if(!empty($result))
    {
      foreach($result as $rs)
      {
        $sheet->setCellValue('A'.$row, $this->products_model->get_barcode($rs->product_code));
        $sheet->setCellValue('B'.$row, $rs->product_code);
        $sheet->setCellValue('C'.$row, $rs->qty);
        $row++;
      }

      $sheet->getStyle('A2:A'.($row -1))->getNumberFormat()->setFormatCode('0');
    }

    setToken($token);
    $file_name = "Report Stock Zone.xlsx";
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); /// form excel 2007 XLSX
    header('Content-Disposition: attachment;filename="'.$file_name.'"');
    $writer = new Xlsx($excel);
    $writer->save('php://output');
  }
}
