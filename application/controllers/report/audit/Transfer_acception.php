<?php
defined('BASEPATH') OR exit('No direct script access allowed');
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class Transfer_acception extends PS_Controller
{
  public $menu_code = 'RADACT';
	public $menu_group_code = 'RE';
  public $menu_sub_group_code = 'REAUDIT';
	public $title = 'รายงานเอกสารโอนคลังที่ต้องกดรับ';  	
	public $limit = 2000;

  public function __construct()
  {
    parent::__construct();
    $this->home = base_url().'report/audit/Transfer_acception';
    $this->load->model('report/audit/transfer_acception_model');
  }

  public function index()
  {
    $this->load->view('report/audit/report_transfer_acception');
  }


  public function get_report()
  {
    $sc = TRUE;

    $json = json_decode($this->input->post('json') ?? '{}');

    if( ! empty($json))
    {
      $docList = array('WW', 'MV', 'RT', 'RN', 'WR', 'SM');

      if($json->allRole == 0 && ! empty($json->role))
      {
        $docList = array();

        foreach($json->role as $role)
        {
          $docList[] = $role;
        }
      }

      $filter = array(
        'from_date' => $json->fromDate,
        'to_date' => $json->toDate,
        'is_accept' => $json->is_accept
      );

      $ds = array();

      $no = 1;

      foreach($docList as $doc)
      {
        $list = $this->transfer_acception_model->get($doc, $filter);

        if( ! empty($list))
        {
          foreach($list as $rs)
          {
            $name = ($doc != 'WW' && $doc != 'MV') ? $rs->accept_name : NULL;
            $owner = ($doc != 'WW' && $doc != 'MV') ? $rs->owner_name : $this->transfer_acception_model->get_owner_list($doc, $rs->code);

            if($rs->is_accept == 1 && empty($rs->accept_by))
            {
              $name = $this->transfer_acception_model->get_accept_list($doc, $rs->code);
            }

            $arr = array(
              'no' => $no,
              'date_add' => thai_date($rs->date_add),
              'code' => $rs->code,
              'is_accept' => $rs->is_accept == 1 ? 'กดรับแล้ว' : 'ยังไม่กดรับ',
              'owner_name' => $owner,
              'accept_by' => $name,
              'accept_on' => thai_date($rs->accept_on, TRUE),
              'accept_remark' => $rs->accept_remark
            );

            array_push($ds, $arr);

            $no++;
          }
        }
      }
    }
    else
    {
      $sc = FALSE;
      $this->error = "Missing required parameter";
    }

    $arr = array(
      'status' => $sc === TRUE ? 'success' : 'failed',
      'message' => $sc === TRUE ? 'success' : $this->error,
      'data' => $sc === TRUE ? $ds : NULL
    );

    echo json_encode($arr);
  }



  public function do_export()
  {
    $token = $this->input->post('token');
		$roles = $this->input->post('role');
		$all  = $this->input->post('allRole') ? TRUE : FALSE;
    $is_accept = $this->input->post('is_accept');
		$fromDate = $this->input->post('fromDate');
    $toDate = $this->input->post('toDate');

    $docList = array('WW', 'MV', 'RT', 'RN', 'WR', 'SM');

    if($all == 0 && ! empty($roles))
    {
      $docList = array();

      foreach($roles as $role)
      {
        $docList[] = $role;
      }
    }

    $filter = array(
      'from_date' => $fromDate,
      'to_date' => $toDate,
      'is_accept' => $is_accept
    );

    $excel = new Spreadsheet();
    $sheet = $excel->getActiveSheet();    
    $sheet->setTitle('Transfer Acception Report');

    //--- set report title header
    $sheet->getColumnDimension('A')->setWidth(8);
    $sheet->getColumnDimension('B')->setWidth(15);
    $sheet->getColumnDimension('C')->setWidth(20);
    $sheet->getColumnDimension('D')->setWidth(30);
    $sheet->getColumnDimension('E')->setWidth(15);
    $sheet->getColumnDimension('F')->setWidth(30);
    $sheet->getColumnDimension('G')->setWidth(25);
    $sheet->getColumnDimension('H')->setWidth(50);

    $sheet->setCellValue('A1', 'รายงานเอกสารโอนคลังที่ต้องกดรับ');

    //--- set Table header
    $sheet->setCellValue('A2', 'ลำดับ');
    $sheet->setCellValue('B2', 'วันที่');
    $sheet->setCellValue('C2', 'เลขที่');
    $sheet->setCellValue('D2', 'เจ้าของโซน');
    $sheet->setCellValue('E2', 'การกดรับ');
    $sheet->setCellValue('F2', 'กดรับโดย');
    $sheet->setCellValue('G2', 'วันที่กดรับ');
    $sheet->setCellValue('H2', 'หมายเหตุ');


		if(! empty($docList))
		{
      $no = 1;
      $row = 3;

      foreach($docList as $doc)
      {
        $list = $this->transfer_acception_model->get($doc, $filter);

        if( ! empty($list))
        {
          foreach($list as $rs)
          {
            $name = ($doc != 'WW' && $doc != 'MV') ? $rs->accept_name : NULL;
            $owner = ($doc != 'WW' && $doc != 'MV') ? $rs->owner_name : $this->transfer_acception_model->get_owner_list($doc, $rs->code);

            if($rs->is_accept == 1 && empty($rs->accept_by))
            {
              $name = $this->transfer_acception_model->get_accept_list($doc, $rs->code);
            }

            $sheet->setCellValue('A'.$row, $no);
  	        $sheet->setCellValue('B'.$row, thai_date($rs->date_add, FALSE, '/'));
  					$sheet->setCellValue('C'.$row, $rs->code);
            $sheet->setCellValue('D'.$row, $owner);
  	        $sheet->setCellValue('E'.$row, $rs->is_accept == 1 ? 'กดรับแล้ว' : 'ยังไม่กดรับ');
  	        $sheet->setCellValue('F'.$row, $name);
  					$sheet->setCellValue('G'.$row, thai_date($rs->accept_on, TRUE, '/'));
            $sheet->setCellValue('H'.$row, $rs->accept_remark);

  					$no++;
  					$row++;
          }
        }
      }
		} //--- end WO

    setToken($token);
    $file_name = "รายงานเอกสารโอนคลังที่ต้องกดรับ ".date('Ymd').".xlsx";
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); /// form excel 2007 XLSX
    header('Content-Disposition: attachment;filename="'.$file_name.'"');
    $writer = new Xlsx($excel);
    $writer->save('php://output');
  }
}
