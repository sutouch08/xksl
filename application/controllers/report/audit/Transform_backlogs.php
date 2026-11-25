<?php
defined('BASEPATH') OR exit('No direct script access allowed');
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class Transform_backlogs extends PS_Controller
{
  public $menu_code = 'RAWQBL';
	public $menu_group_code = 'RE';
  public $menu_sub_group_code = 'REAUDIT';
	public $title = 'รายงานสินค้าแปรสภาพค้างรับ';
  
  public function __construct()
  {
    parent::__construct();
    $this->home = base_url().'report/audit/transform_backlogs';
    $this->load->model('report/audit/transform_backlogs_model');
  }

  public function index()
  {
    $this->load->view('report/audit/report_transform_backlogs');
  }


	public function get_report()
	{
    ini_set('memory_limit','2048M'); // This also needs to be increased in some cases. Can be changed to a higher value as per need)
    ini_set('sqlsrv.ClientBufferMaxKBSize','2097152'); // Setting to 2048M
    ini_set('sqlsrv.client_buffer_max_kb_size','2097152'); // Setting to 512M - for pdo_sqlsrv

		$allUser = $this->input->get('allUser');
		$dname = $this->input->get('dname');
		$allPd = $this->input->get('allProduct');
		$pdFrom = $this->input->get('pdFrom');
		$pdTo = $this->input->get('pdTo');
		$fromDate = empty($this->input->get('fromDate')) ? "" : from_date($this->input->get('fromDate'));
		$toDate = empty($this->input->get('toDate')) ? "" : to_date($this->input->get('toDate'));

		$arr = array(
			'allUser' => $allUser,
			'dname' => $dname,
			'allPd' => $allPd,
			'pdFrom' => $pdFrom,
			'pdTo' => $pdTo,
			'from_date' => $fromDate,
			'to_date' => $toDate
		);

		$ds = array();
		$total_qty = 0;
		$total_return  = 0;
		$total_balance = 0;
		$total_amount = 0;

		$data = $this->transform_backlogs_model->get_data($arr);

		if(!empty($data))
		{
			$no = 1;

			foreach($data as $rs)
			{
				$arr = array(
					'no' => number($no),
					'user_ref' => $rs->user_ref, //--- ผู้เบิก
					'user' => $rs->user_name, //--- ผู้ทำรายการ
          'date_add' => thai_date($rs->date_add, FALSE),
          'due_date' => empty($rs->due_date) ? NULL : thai_date($rs->due_date, FALSE),
					'order_code' => $rs->order_code,
          'original_code' => $rs->original_code,
					'product_code' => $rs->product_code,
					'price' => number($rs->price,2),
					'qty' => number($rs->qty),
					'return' => number($rs->receive),
					'balance' => number($rs->balance),
					'amount' => number($rs->balance * $rs->price, 2)
				);

				array_push($ds, $arr);

				$total_qty += $rs->qty;
				$total_return += $rs->receive;
				$total_balance += $rs->balance;
				$total_amount += ($rs->balance * $rs->price);
				$no++;
			}

			$arr = array(
				'total_qty' => number($total_qty),
				'total_return' => number($total_return),
				'total_balance' => number($total_balance),
				'total_amount' => number($total_amount, 2)
			);

			array_push($ds, $arr);

		}
		else
		{
			$arr = array(
				'total_qty' => number($total_qty),
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
    ini_set('memory_limit','2048M'); // This also needs to be increased in some cases. Can be changed to a higher value as per need)
    ini_set('sqlsrv.ClientBufferMaxKBSize','2097152'); // Setting to 2048M
    ini_set('sqlsrv.client_buffer_max_kb_size','2097152'); // Setting to 512M - for pdo_sqlsrv
    
		$allUser = $this->input->post('allUser');
		$dname = $this->input->post('dname'); //--- orders.user_ref
		$allPd = $this->input->post('allProduct');
		$pdFrom = $this->input->post('pdFrom');
		$pdTo = $this->input->post('pdTo');
		$fromDate = empty($this->input->post('fromDate')) ? "" : from_date($this->input->post('fromDate'));
		$toDate = empty($this->input->post('toDate'))? "" : to_date($this->input->post('toDate'));
		$token = $this->input->post('token');

		//---  Report title
    $report_title = "รายงานสินค้าแปรสภาพค้างร้บ (วันที่พิมพ์รายงาน : ".date('d/m/Y H:i').")";
    $emp_title = 'ผู้เบิก :  '. ($allUser == 1 ? 'ทั้งหมด' : $dname);
    $pd_title = 'สินค้า :  '. ($allPd == 1 ? 'ทั้งหมด' : '('.$pdFrom.') - ('.$pdTo.')');
		$date_title = 'วันที่เอกสาร : '.((empty($fromDate) OR empty($toDate)) ? 'ทั้งหมด' : thai_date($fromDate, '/').' - '.thai_date($toDate, '/'));

		$arr = array(
			'allUser' => $allUser,
			'dname' => $dname, //--- user_ref
			'allPd' => $allPd,
			'pdFrom' => $pdFrom,
			'pdTo' => $pdTo,
			'from_date' => $fromDate,
			'to_date' => $toDate
		);

    $data = $this->transform_backlogs_model->get_data($arr);

    $excel = new Spreadsheet();
    $sheet = $excel->getActiveSheet();

    $sheet->setTitle('WQ-WV ค้างรับ');

    //--- set report title header
    $sheet->setCellValue('A1', $report_title);
    $sheet->mergeCells('A1:L1');
    $sheet->setCellValue('A2', $emp_title);
    $sheet->mergeCells('A2:L2');
    $sheet->setCellValue('A3', $pd_title);
    $sheet->mergeCells('A3:L3');
    $sheet->setCellValue('A4', $date_title);
    $sheet->mergeCells('A4:L4');

    //--- set Table header
		$row = 5;

    $sheet->setCellValue('A'.$row, '#');
    $sheet->setCellValue('B'.$row, 'ผู้เบิก(User)');
    $sheet->setCellValue('C'.$row, 'ผู้ทำรายการ');
    $sheet->setCellValue('D'.$row, 'วันที่');
    $sheet->setCellValue('E'.$row, 'วันที่ต้องการของ');
    $sheet->setCellValue('F'.$row, 'เลขที่เอกสาร');
    $sheet->setCellValue('G'.$row, 'รหัสสินค้า(เบิก)');
    $sheet->setCellValue('H'.$row, 'รหัสสินค้า(รับ)');
    $sheet->setCellValue('I'.$row, 'ราคา');
    $sheet->setCellValue('J'.$row, 'เบิก');
    $sheet->setCellValue('K'.$row, 'รับ');
    $sheet->setCellValue('L'.$row, 'ค้างรับ');
    $sheet->setCellValue('M'.$row, 'มูลค่าคงรับ');

    //---- กำหนดความกว้างของคอลัมภ์
    $sheet->getColumnDimension('B')->setWidth(20);
    $sheet->getColumnDimension('C')->setWidth(20);
    $sheet->getColumnDimension('D')->setWidth(20);
    $sheet->getColumnDimension('E')->setWidth(20);
    $sheet->getColumnDimension('F')->setWidth(20);
    $sheet->getColumnDimension('G')->setWidth(30);
    $sheet->getColumnDimension('H')->setWidth(30);
    $sheet->getColumnDimension('I')->setWidth(15);
    $sheet->getColumnDimension('J')->setWidth(15);
    $sheet->getColumnDimension('K')->setWidth(15);
    $sheet->getColumnDimension('L')->setWidth(15);
    $sheet->getColumnDimension('M')->setWidth(15);

		$row++;


    if(!empty($data))
    {
      $no = 1;

      foreach($data as $rs)
      {
        $sheet->setCellValue('A'.$row, $no);
        $sheet->setCellValue('B'.$row, $rs->user_ref);
        $sheet->setCellValue('C'.$row, $rs->user_name);
        $sheet->setCellValue('D'.$row, thai_date($rs->date_add));
        $sheet->setCellValue('E'.$row, empty($rs->due_date) ? NULL : thai_date($rs->due_date));
        $sheet->setCellValue('F'.$row, $rs->order_code);
        $sheet->setCellValue('G'.$row, $rs->original_code);
        $sheet->setCellValue('H'.$row, $rs->product_code);
        $sheet->setCellValue('I'.$row, $rs->price);
        $sheet->setCellValue('J'.$row, $rs->qty);
        $sheet->setCellValue('K'.$row, $rs->receive);
        $sheet->setCellValue('L'.$row, $rs->balance);
        $sheet->setCellValue('M'.$row, ($rs->balance * $rs->price));

        $no++;
        $row++;
      }

			$re = $row - 1;

		$sheet->setCellValue("A{$row}", 'รวม');
		$sheet->mergeCells("A{$row}:H{$row}");
		$sheet->getStyle("A{$row}")->getAlignment()->setHorizontal('right');

		$sheet->setCellValue("J{$row}", "=SUM(I6:I{$re})");
		$sheet->setCellValue("K{$row}", "=SUM(J6:J{$re})");
		$sheet->setCellValue("L{$row}", "=SUM(K6:K{$re})");
		$sheet->setCellValue("M{$row}", "=SUM(L6:L{$re})");

		$sheet->getStyle("I6:I{$row}")->getNumberFormat()->setFormatCode('#,##0.00');
		$sheet->getStyle("J6:L{$row}")->getNumberFormat()->setFormatCode('#,##0');
		$sheet->getStyle("M6:M{$row}")->getNumberFormat()->setFormatCode('#,##0.00');
    }
		else
		{
			$sheet->setCellValue('A'.$row, "ไม่พบข้อมูลตามเงื่อนไขที่กำหนด");
			$row++;
		}

    setToken($token);
    $file_name = "รายงานสินค้าแปรสภาพค้างรับ_".date('dmY').".xlsx";
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); /// form excel 2007 XLSX
    header('Content-Disposition: attachment;filename="'.$file_name.'"');
    $writer = new Xlsx($excel);
    $writer->save('php://output');
  }
}
