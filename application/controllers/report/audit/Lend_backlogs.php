<?php
defined('BASEPATH') OR exit('No direct script access allowed');
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class Lend_backlogs extends PS_Controller
{
  public $menu_code = 'RALNBL';
	public $menu_group_code = 'RE';
  public $menu_sub_group_code = 'REAUDIT';
	public $title = 'รายงานสินค้ายืมยังไม่คืน';
  
  public function __construct()
  {
    parent::__construct();
    $this->home = base_url().'report/audit/lend_backlogs';
    $this->load->model('report/audit/lend_backlogs_model');
    $this->load->model('masters/products_model');
  }

  public function index()
  {
    $this->load->view('report/audit/report_lend_backlogs');
  }


	public function get_report()
	{
		$allEmp = $this->input->get('allEmp');
		$empId = $this->input->get('empId');
		$allPd = $this->input->get('allProduct');
		$pdFrom = $this->input->get('pdFrom');
		$pdTo = $this->input->get('pdTo');
    $from_date = $this->input->get('fromDate');
    $to_date = $this->input->get('toDate');
		$fromDate = empty($from_date) ? '2021-01-01' : $from_date;
		$toDate = empty($to_date) ? now() : $to_date;

		$arr = array(
			'allEmp' => $allEmp,
			'empId' => $empId,
			'allPd' => $allPd,
			'pdFrom' => $pdFrom,
			'pdTo' => $pdTo,
			'from_date' => $fromDate,
			'to_date' => $toDate
		);

		$ds = array();
		$total_lend = 0;
		$total_return  = 0;
		$total_balance = 0;
		$total_amount = 0;

		$data = $this->lend_backlogs_model->get_data($arr);

		if(!empty($data))
		{
			$no = 1;

			foreach($data as $rs)
			{
				$arr = array(
					'no' => number($no),
					'emp_name' => $rs->empName, //--- ผู้เบิก
					'user_ref' => $rs->user_ref, //--- ผู้รับ
					'user' => $rs->user_name, //--- ผู้ทำรายการ
					'order_code' => $rs->order_code,
					'product_code' => $rs->product_code,
					'price' => number($rs->price,2),
					'lend' => number($rs->qty),
					'return' => number($rs->receive),
					'balance' => number($rs->balance),
					'amount' => number($rs->balance * $rs->price, 2)
				);

				array_push($ds, $arr);

				$total_lend += $rs->qty;
				$total_return += $rs->receive;
				$total_balance += $rs->balance;
				$total_amount += ($rs->balance * $rs->price);
				$no++;
			}

			$arr = array(
				'total_lend' => number($total_lend),
				'total_return' => number($total_return),
				'total_balance' => number($total_balance),
				'total_amount' => number($total_amount, 2)
			);

			array_push($ds, $arr);

		}
		else
		{
			$arr = array(
				'total_lend' => number($total_lend),
				'total_return' => number($total_return),
				'total_balance' => number($total_balance),
				'total_amount' => number($total_amount, 2)
			);

			array_push($ds, $arr);
		}

		echo json_encode($ds);
	}


  public function do_export()
  {
    ini_set('memory_limit','2048M');
		ini_set('max_execution_time', 600);
    ini_set('sqlsrv.ClientBufferMaxKBSize','2097152');
    ini_set('sqlsrv.client_buffer_max_kb_size','2097152');
    
		$allEmp = $this->input->post('allEmp');
		$empId = $this->input->post('empId');
		$empName = $this->input->post('empName');
		$allPd = $this->input->post('allProduct');
		$pdFrom = $this->input->post('pdFrom');
		$pdTo = $this->input->post('pdTo');
    $from_date = $this->input->post('fromDate');
    $to_date = $this->input->post('toDate');
		$fromDate = empty($from_date) ? '2021-01-01' : $from_date;
		$toDate = empty($to_date) ? now() : $to_date;
		$token = $this->input->post('token');

		//---  Report title
    $report_title = "รายงานสินค้ายืมยังไม่คืน (วันที่พิมพ์รายงาน : ".date('d/m/Y H:i').")";
    $emp_title = 'ผู้ยืม :  '. ($allEmp == 1 ? 'ทั้งหมด' : $empName);
    $pd_title = 'สินค้า :  '. ($allPd == 1 ? 'ทั้งหมด' : '('.$pdFrom.') - ('.$pdTo.')');
		$date_title = 'วันที่เอกสาร : '.thai_date($fromDate, '/').' - '.thai_date($toDate, '/');

		$arr = array(
			'allEmp' => $allEmp,
			'empId' => $empId,
			'allPd' => $allPd,
			'pdFrom' => $pdFrom,
			'pdTo' => $pdTo,
			'from_date' => $fromDate,
			'to_date' => $toDate
		);

    $data = $this->lend_backlogs_model->get_data($arr);

    $excel = new Spreadsheet();
		$sheet = $excel->getActiveSheet();
		$sheet->setTitle('Lend backlogs Report');

    //--- set report title header
    $sheet->setCellValue('A1', $report_title);
    $sheet->mergeCells('A1:K1');
    $sheet->setCellValue('A2', $emp_title);
    $sheet->mergeCells('A2:K2');
    $sheet->setCellValue('A3', $pd_title);
    $sheet->mergeCells('A3:K3');
    $sheet->setCellValue('A4', $date_title);
    $sheet->mergeCells('A4:K4');

    //--- set Table header
		$row = 5;

    $sheet->setCellValue('A'.$row, '#');
    $sheet->setCellValue('B'.$row, 'ผู้ยืม');
    $sheet->setCellValue('C'.$row, 'ผู้รับ');
    $sheet->setCellValue('D'.$row, 'ผู้ทำรายการ');
    $sheet->setCellValue('E'.$row, 'เลขที่เอกสาร');
    $sheet->setCellValue('F'.$row, 'รหัสสินค้า');
    $sheet->setCellValue('G'.$row, 'ราคา');
    $sheet->setCellValue('H'.$row, 'ยืม');
    $sheet->setCellValue('I'.$row, 'คืน');
    $sheet->setCellValue('J'.$row, 'คงค้าง');
    $sheet->setCellValue('K'.$row, 'มูลค่าคงค้าง');

    //---- กำหนดความกว้างของคอลัมภ์
    $sheet->getColumnDimension('B')->setWidth(20);
    $sheet->getColumnDimension('C')->setWidth(20);
    $sheet->getColumnDimension('D')->setWidth(20);
    $sheet->getColumnDimension('E')->setWidth(20);
    $sheet->getColumnDimension('F')->setWidth(30);
    $sheet->getColumnDimension('G')->setWidth(15);
    $sheet->getColumnDimension('H')->setWidth(15);
    $sheet->getColumnDimension('I')->setWidth(15);
    $sheet->getColumnDimension('J')->setWidth(15);
    $sheet->getColumnDimension('K')->setWidth(15);

		$row++;


    if(!empty($data))
    {
      $no = 1;

      foreach($data as $rs)
      {
        $sheet->setCellValue('A'.$row, $no);
        $sheet->setCellValue('B'.$row, $rs->empName);
        $sheet->setCellValue('C'.$row, $rs->user_ref);
        $sheet->setCellValue('D'.$row, $rs->user_name);
        $sheet->setCellValue('E'.$row, $rs->order_code);
        $sheet->setCellValue('F'.$row, $rs->product_code);
        $sheet->setCellValue('G'.$row, $rs->price);
        $sheet->setCellValue('H'.$row, $rs->qty);
        $sheet->setCellValue('I'.$row, $rs->receive);
        $sheet->setCellValue('J'.$row, $rs->balance);
        $sheet->setCellValue('K'.$row, ($rs->balance * $rs->price));

        $no++;
        $row++;
      }
    }

		$re = $row - 1;

		$sheet->setCellValue("A{$row}", 'รวม');
		$sheet->mergeCells("A{$row}:G{$row}");
		$sheet->getStyle("A{$row}")->getAlignment()->setHorizontal('right');

		$sheet->setCellValue("H{$row}", "=SUM(H6:H{$re})");
		$sheet->setCellValue("I{$row}", "=SUM(I6:I{$re})");
		$sheet->setCellValue("J{$row}", "=SUM(J6:J{$re})");
		$sheet->setCellValue("K{$row}", "=SUM(K6:K{$re})");

    setToken($token);
    $file_name = "Lend_backlogs_".date('dmY').".xlsx";
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); /// form excel 2007 XLSX
    header('Content-Disposition: attachment;filename="'.$file_name.'"');
    $writer = new Xlsx($excel);
    $writer->save('php://output');
  }
}
