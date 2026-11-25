<?php
defined('BASEPATH') OR exit('No direct script access allowed');
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class Stock extends PS_Controller
{
  public $menu_code = 'ICCKST';
	public $menu_group_code = 'IC';
  public $menu_sub_group_code = 'CHECK';
	public $title = 'ตรวจสอบสต็อกคงเหลือ(SAP)';

  public function __construct()
  {
    parent::__construct();
    $this->home = base_url().'inventory/stock';
    $this->load->model('inventory/sap_stock_model');
  }


  public function index()
  {
    $filter = array(
      'item_code' => get_filter('item_code', 'item_code', ''),
      'zone_code' => get_filter('zone_code', 'zone_code', ''),
      'show_system' => get_filter('show_system', 'show_system', 'no')
    );

		//--- แสดงผลกี่รายการต่อหน้า
		$perpage = get_rows();
		//--- หาก user กำหนดการแสดงผลมามากเกินไป จำกัดไว้แค่ 300
		if($perpage > 300)
		{
			$perpage = 20;
		}

		$segment  = 4; //-- url segment
		$rows     = $this->sap_stock_model->count_rows($filter);
		//--- ส่งตัวแปรเข้าไป 4 ตัว base_url ,  total_row , perpage = 20, segment = 3
		$init	    = pagination_config($this->home.'/index/', $rows, $perpage, $segment);
		$ds   = $this->sap_stock_model->get_list($filter, $perpage, $this->uri->segment($segment));

    $filter['data'] = $ds;

		$this->pagination->initialize($init);
    $this->load->view('inventory/stock/stock_view', $filter);
  }



  public function export()
  {
    $arr = array(
      'item_code' => $this->input->post('item'),
      'zone_code' => $this->input->post('zone'),
      'show_system' => $this->input->post('system')
    );

    $token = $this->input->post('token');
    $data = $this->sap_stock_model->get_list($arr);

    if(!empty($data))
    {
      $excel = new Spreadsheet();
      $sheet = $excel->getActiveSheet();
      $sheet->setTitle('Stock Zone (SAP)');

      $sheet->setCellValue('A1', 'No.');
      $sheet->setCellValue('B1', 'ItemCode');
      $sheet->setCellValue('C1', 'OldCode');
      $sheet->setCellValue('D1', 'Description');
      $sheet->setCellValue('E1', 'BinCode');
      $sheet->setCellValue('F1', 'Bin Description');
      $sheet->setCellValue('G1', 'Qty');

      $no = 1;
      $row = 2;
      foreach($data as $rs)
      {
        $sheet->setCellValue('A'.$row, $no);
        $sheet->setCellValue('B'.$row, $rs->ItemCode);
        if(!empty($rs->U_OLDCODE))
        {
          $sheet->setCellValue('C'.$row, $rs->U_OLDCODE);
        }

        $sheet->setCellValue('D'.$row, $rs->ItemName);
        $sheet->setCellValue('E'.$row, $rs->BinCode);
        $sheet->setCellValue('F'.$row, $rs->Descr);
        $sheet->setCellValue('G'.$row, $rs->OnHandQty);
        $no++;
        $row++;
      }
    }

    setToken($token);
    $file_name = "StockZone(SAP).xlsx";
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); /// form excel 2007 XLSX
    header('Content-Disposition: attachment;filename="'.$file_name.'"');
    $writer = new Xlsx($excel);
    $writer->save('php://output');
  }

  function clear_filter(){
    $filter = array('item_code', 'zone_code', 'show_system');
    clear_filter($filter);
    echo 'done';
  }
}
