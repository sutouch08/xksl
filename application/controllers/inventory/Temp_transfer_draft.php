<?php
defined('BASEPATH') OR exit('No direct script access allowed');
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class Temp_transfer_draft extends PS_Controller
{
  public $menu_code = 'TETDCK';
	public $menu_group_code = 'TE';
  public $menu_sub_group_code = 'TETRANSFER';
	public $title = 'ตรวจสอบ Transfer Draft';
  
  public function __construct()
  {
    parent::__construct();
    $this->home = base_url().'inventory/temp_transfer_draft';
    $this->load->model('inventory/temp_transfer_draft_model');
  }


  public function index()
  {
    $filter = array(
      'code'          => get_filter('code', 'temp_tr_code', ''),
      'customer'      => get_filter('customer', 'temp_tr_customer', ''),
      'from_date'     => get_filter('from_date', 'temp_tr_from_date', ''),
      'to_date'       => get_filter('to_date', 'temp_tr_to_date', ''),
      'status'      => get_filter('status', 'temp_tr_status', 'all'),
      'is_received' => get_filter('is_received', 'temp_tr_is_received', 'all')
    );

		//--- แสดงผลกี่รายการต่อหน้า
		$perpage = get_rows();		
		$segment  = 4; //-- url segment
		$rows     = $this->temp_transfer_draft_model->count_rows($filter, 8);
		
		$init	    = pagination_config($this->home.'/index/', $rows, $perpage, $segment);
		$orders   = $this->temp_transfer_draft_model->get_list($filter, $perpage, $this->uri->segment($segment), 8);

    $filter['orders'] = $orders;

		$this->pagination->initialize($init);
    $this->load->view('inventory/temp_transfer_draft/temp_list', $filter);
  }



  public function get_detail($code)
  {
    $this->load->model('stock/stock_model');
    $detail = $this->temp_transfer_draft_model->get_detail($code);
    if(!empty($detail))
    {
      foreach($detail as $rs)
      {
        $rs->onhand = $this->stock_model->get_stock_zone($rs->F_FROM_BIN, $rs->ItemCode);
      }
    }

    $ds['details'] = $detail;
    $ds['code'] = $code;
    $this->load->view('inventory/temp_transfer_draft/temp_detail', $ds);
  }



  public function export_diff()
  {
    $this->load->model('stock/stock_model');
    $token = $this->input->post('token');

    //---  Report title
    $report_title = 'ยอดต่างรายการสินค้าในโซน';
    $channels_title = 'ยอดต่าง = ยอดในโซน - ยอดที่ออเดอร์';
    $pd_title = 'ยอดติดลบคือยอดที่มีในโซนน้อยกว่าที่สั่งมา';

    $excel = new Spreadsheet();
    $sheet = $excel->getActiveSheet();
    $sheet->setTitle('Item Diff');

    //--- set report title header
    $sheet->setCellValue('A1', $report_title);
    $sheet->mergeCells('A1:G1');
    $sheet->setCellValue('A2', $channels_title);
    $sheet->mergeCells('A2:G2');
    $sheet->setCellValue('A3', $pd_title);
    $sheet->mergeCells('A3:G3');
    $sheet->setCellValue('A4', 'วันที่เอกสาร : '.thai_date(now()));
    $sheet->mergeCells('A4:G4');

    $sheet->setCellValue('A5', 'ลำดับ');
    $sheet->setCellValue('B5', 'BinCode');
    $sheet->setCellValue('C5', 'ItemCode');
    $sheet->setCellValue('D5', 'onhand - order');

    $ds = array();
    $orders = $this->temp_transfer_draft_model->get_error_list();
    if(!empty($orders))
    {
      foreach($orders as $order)
      {
        $details = $this->temp_transfer_draft_model->get_detail($order->code);
        if(!empty($details))
        {
          foreach($details as $rs)
          {
            $onhand = $this->stock_model->get_stock_zone($rs->F_FROM_BIN, $rs->ItemCode);
            if($rs->Quantity > $onhand)
            {
              $diff = $onhand - $rs->Quantity;
              if(isset($ds[$rs->F_FROM_BIN][$rs->ItemCode]))
              {
                $ds[$rs->F_FROM_BIN][$rs->ItemCode] += $diff;
              }
              else
              {
                $ds[$rs->F_FROM_BIN][$rs->ItemCode] = $diff;
              }
            }
          }
        }
      }

      $row = 6;

      if(!empty($ds))
      {
        $no = 1;
        foreach($ds as $bin => $items)
        {
          foreach($items as $item => $qty)
          {
            $sheet->setCellValue('A'.$row, $no);

            $sheet->setCellValue('B'.$row, $bin);

            $sheet->setCellValue('C'.$row, $item);

            $sheet->setCellValue('D'.$row, $qty);

            $no++;
            $row++;
          }
        }
      }
    }

    setToken($token);
    $file_name = "Transfer draft stock diff.xlsx";
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); /// form excel 2007 XLSX
    header('Content-Disposition: attachment;filename="'.$file_name.'"');
    $writer = new Xlsx($excel);
    $writer->save('php://output');
  }


  public function clear_filter()
  {
    $filter = array(
      'temp_tr_code',
      'temp_tr_supplier',
      'temp_tr_from_date',
      'temp_tr_to_date',
      'temp_tr_status',
      'temp_tr_is_received'
    );

    clear_filter($filter);
  }
}
