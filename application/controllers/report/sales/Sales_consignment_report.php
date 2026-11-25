<?php
defined('BASEPATH') OR exit('No direct script access allowed');
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class Sales_consignment_report extends PS_Controller
{
  public $menu_code = 'RSOCMD';
	public $menu_group_code = 'RE';
  public $menu_sub_group_code = 'RESALE';
	public $title = 'รายงานตัดยอดฝากขายเทียม แสดงรายการสินค้า';
  public $filter;
  public $error;
  public function __construct()
  {
    parent::__construct();
    $this->home = base_url().'report/sales/sales_consignment_report';
    $this->load->model('report/sales/sales_consignment_report_model');
    $this->load->model('masters/products_model');
    $this->load->model('masters/customers_model');
  }

  public function index()
  {
    $this->load->model('masters/warehouse_model');
    $whList = $this->warehouse_model->get_consignment_list();
    $ds['whList'] = $whList;
    $this->load->view('report/sales/sales_consignment_report', $ds);
  }


  public function get_report()
  {
    ini_set('memory_limit','2048M'); // This also needs to be increased in some cases. Can be changed to a higher value as per need)
    ini_set('sqlsrv.ClientBufferMaxKBSize','2097152'); // Setting to 2048M
    ini_set('sqlsrv.client_buffer_max_kb_size','2097152'); // Setting to 512M - for pdo_sqlsrv

    $sc = TRUE;
    $allProduct = $this->input->get('allProduct');
    $pdFrom = $this->input->get('pdFrom');
    $pdTo = $this->input->get('pdTo');

    $allCustomer = $this->input->get('allCustomer');
    $cusFrom = $this->input->get('cusFrom');
    $cusTo = $this->input->get('cusTo');

    $allWhouse = $this->input->get('allWhouse');
    $warehouse = $this->input->get('warehouse');

    $allZone = $this->input->get('allZone');
    $zoneCode = $this->input->get('zoneCode');
    $zoneName = $this->input->get('zoneName');

    $fromDate = $this->input->get('fromDate');
    $toDate = $this->input->get('toDate');

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

    $bs = array();

    $filter = array(
      'fromDate' => $fromDate,
      'toDate' => $toDate,
      'allProduct' => $allProduct,
      'pdFrom' => $pdFrom,
      'pdTo' => $pdTo,
      'allCustomer' => $allCustomer,
      'cusFrom' => $cusFrom,
      'cusTo' => $cusTo,
      'allWarehouse' => $allWhouse,
      'warehouse_code' => $warehouse,
      'allZone' => $allZone,
      'zone_code' => $zoneCode
    );

    $result = $this->sales_consignment_report_model->get_data($filter);

    if(! empty($result))
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
        $totalDiscount = 0;
        $totalAmount = 0;
        $totalCost = 0;

        foreach($result as $rs)
        {
          $arr = array(
            'no' => number($no),
            'date_add' => thai_date($rs->date_add, FALSE),
            'reference' => $rs->reference,
            'product_code' => $rs->product_code,
            'product_name' => $rs->product_name,
            'cost' => number($rs->cost, 2),
            'price' => number($rs->price, 2),
            'discount_label' => $rs->discount_label,
            'sell' => number($rs->sell, 2),
            'qty' => number($rs->qty),
            'total_discount' => number(($rs->discount_amount * $rs->qty), 2),
            'total_amount' => number($rs->total_amount, 2),
            'total_cost' => number($rs->total_cost, 2),
            'customer_code' => $rs->customer_code,
            'customer_name' => $rs->customer_name,
            'warehouse_code' => $rs->warehouse_code,
            'warehouse_name' => $rs->warehouse_name,
            'zone_code' => $rs->zone_code,
            'zone_name' => $rs->zone_name
          );

          array_push($bs, $arr);
          $totalQty += $rs->qty;
          $totalDiscount += ($rs->qty * $rs->discount_amount);
          $totalAmount += $rs->total_amount;
          $totalCost += $rs->total_cost;
          $no++;
        }

        $arr = array(
          'totalQty' => number($totalQty),
          'totalDiscount' => number($totalDiscount, 2),
          'totalAmount' => number($totalAmount, 2),
          'totalCost' => number($totalCost, 2)
        );

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
    ini_set('memory_limit','2048M'); // This also needs to be increased in some cases. Can be changed to a higher value as per need)
    ini_set('sqlsrv.ClientBufferMaxKBSize','2097152'); // Setting to 2048M
    ini_set('sqlsrv.client_buffer_max_kb_size','2097152'); // Setting to 512M - for pdo_sqlsrv
    
    $sc = TRUE;
    $allProduct = $this->input->post('allProduct');
    $pdFrom = $this->input->post('pdFrom');
    $pdTo = $this->input->post('pdTo');

    $allCustomer = $this->input->post('allCustomer');
    $cusFrom = $this->input->post('cusFrom');
    $cusTo = $this->input->post('cusTo');

    $allWhouse = $this->input->post('allWhouse');
    $warehouse = $this->input->post('warehouse');

    $allZone = $this->input->post('allZone');
    $zoneCode = $this->input->post('zoneCode');
    $zoneName = $this->input->post('zoneName');

    $fromDate = $this->input->post('fromDate');
    $toDate = $this->input->post('toDate');

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
    $report_title = "รายงานตัดยอดฝากขายเทียม";
    $whList = $allWhouse == 1 ? 'ทั้งหมด' : $wh_list;
    $zoneList = $allZone == 1 ? 'ทั้งหมด' : $zoneCode." - ".$zoneName;
    $productList  = $allProduct == 1 ? 'ทั้งหมด' : '('.$pdFrom.') - ('.$pdTo.')';
    $dateRange = "วันที่ ".thai_date($fromDate, FALSE, '/').' - '.thai_date($toDate, FALSE, '/');
    $cusList = $allCustomer == 1 ? 'ทั้งหมด' : "({$cusFrom}) - ({$cusTo})";


    $filter = array(
      'fromDate' => $fromDate,
      'toDate' => $toDate,
      'allProduct' => $allProduct,
      'pdFrom' => $pdFrom,
      'pdTo' => $pdTo,
      'allCustomer' => $allCustomer,
      'cusFrom' => $cusFrom,
      'cusTo' => $cusTo,
      'allWarehouse' => $allWhouse,
      'warehouse_code' => $warehouse,
      'allZone' => $allZone,
      'zone_code' => $zoneCode
    );

    $result = $this->sales_consignment_report_model->get_data($filter);

    $excel = new Spreadsheet();
    $sheet = $excel->getActiveSheet();    
    $sheet->setTitle('Sales Consignment Report (WD)');

    //--- set report title header
    $sheet->setCellValue('A1', $report_title);
    $sheet->setCellValue('A2', $dateRange);
    $sheet->setCellValue('A3', 'ลูกค้า : '.$cusList);
    $sheet->setCellValue('A4', 'คลัง : '.$whList);
    $sheet->setCellValue('A5', 'โซน : '.$zoneList);
    $sheet->setCellValue('A6', 'สินค้า : '.$productList);

    $row = 6;
    //--- set Table header
    $sheet->setCellValue("A{$row}", "วันที่");
    $sheet->setCellValue("B{$row}", "เลขที่");
    $sheet->setCellValue("C{$row}", "รหัส");
    $sheet->setCellValue("D{$row}", "สินค้า");
    $sheet->setCellValue("E{$row}", "ทุน");
    $sheet->setCellValue("F{$row}", "ราคา");
    $sheet->setCellValue("G{$row}", "ขาย");
    $sheet->setCellValue("H{$row}", "จำนวน");
    $sheet->setCellValue("I{$row}", "ส่วนลด");
    $sheet->setCellValue("J{$row}", "มูลค่าส่วนลด");
    $sheet->setCellValue("K{$row}", "มูลค่ารวม");
    $sheet->setCellValue("L{$row}", "ทุนรวม");
    $sheet->setCellValue("M{$row}", "รหัสลูกค้า");
    $sheet->setCellValue("N{$row}", "ชื่อลูกค้า");
    $sheet->setCellValue("O{$row}", "รหัสคลัง");
    $sheet->setCellValue("P{$row}", "ชื่อคลัง");
    $sheet->setCellValue("Q{$row}", "รหัสโซน");
    $sheet->setCellValue("R{$row}", "ชื่อโซน");

    $row++;

    if(! empty($result))
    {
      // print_r($result);
      foreach($result as $rs)
      {
        $sheet->setCellValue("A{$row}", thai_date($rs->date_add, FALSE, '/'));
        $sheet->setCellValue("B{$row}", $rs->reference);
        $sheet->setCellValue("C{$row}", $rs->product_code);
        $sheet->setCellValue("D{$row}", $rs->product_name);
        $sheet->setCellValue("E{$row}", $rs->cost);
        $sheet->setCellValue("F{$row}", $rs->price);
        $sheet->setCellValue("G{$row}", $rs->sell);
        $sheet->setCellValue("H{$row}", $rs->qty);
        $sheet->setCellValueExplicit("I{$row}", $rs->discount_label, DataType::TYPE_STRING);
        $sheet->setCellValue("J{$row}", $rs->discount_amount);
        $sheet->setCellValue("K{$row}", $rs->total_amount);
        $sheet->setCellValue("L{$row}", $rs->total_cost);
        $sheet->setCellValue("M{$row}", $rs->customer_code);
        $sheet->setCellValue("N{$row}", $rs->customer_name);
        $sheet->setCellValue("O{$row}", $rs->warehouse_code);
        $sheet->setCellValue("P{$row}", $rs->warehouse_name);
        $sheet->setCellValue("Q{$row}", $rs->zone_code);
        $sheet->setCellValue("R{$row}", $rs->zone_name);
        $row++;
      }

      $sheet->getStyle("E7:L{$row}")->getAlignment()->setHorizontal('right');
    }

    setToken($token);
    $file_name = "Report Sales Consignment.xlsx";
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); /// form excel 2007 XLSX
    header('Content-Disposition: attachment;filename="'.$file_name.'"');
    $writer = new Xlsx($excel);
    $writer->save('php://output');
  }
}
