<?php
defined('BASEPATH') OR exit('No direct script access allowed');
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class Stock_balance extends PS_Controller
{
  public $menu_code = 'RICSTB';
	public $menu_group_code = 'RE';
  public $menu_sub_group_code = 'REINVT';
	public $title = 'รายงานสินค้าคงเหลือ';
  
  public function __construct()
  {
    parent::__construct();
    $this->home = base_url().'report/inventory/stock_balance';
    $this->load->model('report/inventory/inventory_report_model');
    $this->load->model('masters/products_model');
  }

  public function index()
  {
    $this->load->model('masters/warehouse_model');
    $whList = $this->warehouse_model->get_all_warehouse();
    $ds['whList'] = $whList;
    $this->load->view('report/inventory/report_stock_balance', $ds);
  }


  public function get_report()
  {
    $allProduct = $this->input->get('allProduct');
    $pdFrom = $this->input->get('pdFrom');
    $pdTo = $this->input->get('pdTo');

    $allWhouse = $this->input->get('allWhouse');
    $warehouse = $this->input->get('warehouse');

    $wh_list = '';
    if(!empty($warehouse))
    {
      $i = 1;
      foreach($warehouse as $wh)
      {
        $wh_list .= $i === 1 ? $wh : ', '.$wh;
        $i++;
      }
    }

    //---  Report title
    $sc['reportDate'] = thai_date(date('Y-m-d'),FALSE, '/');
    $sc['whList']   = $allWhouse == 1 ? 'ทั้งหมด' : $wh_list;
    $sc['productList']   = $allProduct == 1 ? 'ทั้งหมด' : '('.$pdFrom.') - ('.$pdTo.')';


    $result = $this->inventory_report_model->get_current_stock_balance($allProduct, $pdFrom, $pdTo, $allWhouse, $warehouse);

    $bs = array();

    if(!empty($result))
    {
      $no = 1;
      $totalQty = 0;
      $totalAmount = 0;
      foreach($result as $rs)
      {
        $item = $this->products_model->get_item($rs->product_code);
        if(!empty($item))
        {
          $arr = array(
            'no' => number($no),
            'barcode' => $item->barcode,
            'pdCode' => $item->code,
            'pdName' => $item->name,
            'cost' => number($item->cost, 2),
            'qty' => number($rs->qty),
            'amount' => number($item->cost * $rs->qty, 2)
          );

          array_push($bs, $arr);
          $no++;

          $totalQty += $rs->qty;
          $totalAmount += ($rs->qty * $item->cost);
        }

      }

      $arr = array(
        'totalQty' => number($totalQty),
        'totalAmount' => number($totalAmount, 2)
      );

      array_push($bs, $arr);

      $bs;
    }
    else
    {
      $arr = array('nodata' => 'nodata');
      array_push($bs, $arr);
    }

    $sc['bs'] = $bs;

    echo json_encode($sc);
  }


  public function do_export()
  {
    $allProduct = $this->input->post('allProduct');
    $pdFrom = $this->input->post('pdFrom');
    $pdTo = $this->input->post('pdTo');

    $allWhouse = $this->input->post('allWhouse');
    $warehouse = $this->input->post('warehouse');


    $wh_list = '';
    if(!empty($warehouse))
    {
      $i = 1;
      foreach($warehouse as $wh)
      {
        $wh_list .= $i === 1 ? $wh : ', '.$wh;
        $i++;
      }
    }


    //---  Report title
    $report_title = 'รายงานสินค้าคงเหลือ ณ วันที่  '.thai_date(date('Y-m-d'), '/');
    $wh_title     = 'คลัง :  '. ($allWhouse == 1 ? 'ทั้งหมด' : $wh_list);
    $pd_title     = 'สินค้า :  '. ($allProduct == 1 ? 'ทั้งหมด' : '('.$pdFrom.') - ('.$pdTo.')');

    $result = $this->inventory_report_model->get_current_stock_balance($allProduct, $pdFrom, $pdTo, $allWhouse, $warehouse);
    $excel = new Spreadsheet();
    $sheet = $excel->getActiveSheet();
    $sheet->setTitle('Stock Balance Report');

    //--- set report title header
    $sheet->setCellValue('A1', $report_title);
    $sheet->mergeCells('A1:G1');
    $sheet->setCellValue('A2', $wh_title);
    $sheet->mergeCells('A2:G2');
    $sheet->setCellValue('A3', $pd_title);
    $sheet->mergeCells('A3:G3');

    //--- set Table header
    $sheet->setCellValue('A4', 'ลำดับ');
    $sheet->setCellValue('B4', 'บาร์โค้ด');
    $sheet->setCellValue('C4', 'รหัส');
    $sheet->setCellValue('D4', 'สินค้า');
    $sheet->setCellValue('E4', 'ทุน');
    $sheet->setCellValue('F4', 'จำนวน');
    $sheet->setCellValue('G4', 'มูลค่า');

    $row = 5;
    if(!empty($result))
    {
      $no = 1;
      foreach($result as $rs)
      {
        $item = $this->products_model->get_item($rs->product_code);
        if(!empty($item))
        {
          $sheet->setCellValue('A'.$row, $no);
          $sheet->setCellValue('B'.$row, $item->barcode);
          $sheet->setCellValue('C'.$row, $item->code);
          $sheet->setCellValue('D'.$row, $item->name);
          $sheet->setCellValue('E'.$row, $item->cost);
          $sheet->setCellValue('F'.$row, $rs->qty);
          $sheet->setCellValue('G'.$row, ($item->cost * $rs->qty));
          $no++;
          $row++;
        }
      }

      $res = $row -1;

      $sheet->setCellValue('A'.$row, 'รวม');
      $sheet->mergeCells('A'.$row.':E'.$row);
      $sheet->setCellValue('F'.$row, '=SUM(F5:F'.$res.')');
      $sheet->setCellValue('G'.$row, '=SUM(G5:G'.$res.')');

      $sheet->getStyle('A'.$row)->getAlignment()->setHorizontal('right');
      $sheet->getStyle('B5:B'.$res)->getNumberFormat()->setFormatCode('0');
      $sheet->getStyle('F5:G'.$row)->getAlignment()->setHorizontal('right');
      $sheet->getStyle('F5:F'.$row)->getNumberFormat()->setFormatCode('#,##0');
      $sheet->getStyle('G5:G'.$row)->getNumberFormat()->setFormatCode('#,##0.00');
    }


    $file_name = "Report Stock Balance.xlsx";
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); /// form excel 2007 XLSX
    header('Content-Disposition: attachment;filename="'.$file_name.'"');
    $writer = new Xlsx($excel);
    $writer->save('php://output');
  }
}
