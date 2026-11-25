<?php
defined('BASEPATH') or exit('No direct script access allowed');

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class Document_running extends PS_Controller
{
	public $menu_code = 'RADRUN';
	public $menu_group_code = 'RE';
	public $menu_sub_group_code = 'REAUDIT';
	public $title = 'รายงานทะเบียนคุมเอกสารแยกตามประเภท';
	public $limit = 2000;

	public function __construct()
	{
		parent::__construct();
		$this->home = base_url() . 'report/audit/document_running';
		$this->load->model('report/audit/document_model');
	}

	public function index()
	{
		$this->load->view('report/audit/document_running');
	}

	public function do_export()
	{
		ini_set('memory_limit', '2048M'); // This also needs to be increased in some cases. Can be changed to a higher value as per need)
		ini_set('sqlsrv.ClientBufferMaxKBSize', '2097152'); // Setting to 2048M
		ini_set('sqlsrv.client_buffer_max_kb_size', '2097152'); // Setting to 512M - for pdo_sqlsrv

		$token = $this->input->post('token');
		$role = $this->input->post('role');
		$all  = $this->input->post('allRole') ? TRUE : FALSE;
		$fromDate = $this->input->post('fromDate');
		$toDate = $this->input->post('toDate');

		$excel = new Spreadsheet();
		$index = 0;

		if (!empty($role['WO']) or $all)
		{
			$sheet = $excel->getActiveSheet();
			$sheet->setTitle('WO');

			$index++;

			//--- set report title header
			$sheet->getColumnDimension('A')->setWidth(8);
			$sheet->getColumnDimension('B')->setWidth(15);
			$sheet->getColumnDimension('C')->setWidth(15);
			$sheet->getColumnDimension('D')->setWidth(15);
			$sheet->getColumnDimension('E')->setWidth(15);
			$sheet->getColumnDimension('F')->setWidth(15);
			$sheet->getColumnDimension('G')->setWidth(20);

			//--- set Table header
			$sheet->setCellValue('A1', 'ลำดับ');
			$sheet->setCellValue('B1', 'วันที่');
			$sheet->setCellValue('C1', 'เลขที่');
			$sheet->setCellValue('D1', 'SAP DO');
			$sheet->setCellValue('E1', 'ช่องทางขาย');
			$sheet->setCellValue('F1', 'สถานะ');
			$sheet->setCellValue('G1', 'หมายเหตุ');

			$data = $this->document_model->getOrder('S', $fromDate, $toDate);

			if (!empty($data))
			{
				$no = 1;
				$row = 2;

				foreach ($data as $rs)
				{
					$sheet->setCellValue('A' . $row, $no);
					$sheet->setCellValue('B' . $row, thai_date($rs->date_add, FALSE, '/'));
					$sheet->setCellValue('C' . $row, $rs->code);
					$sheet->setCellValue('D' . $row, $rs->inv_code);
					$sheet->setCellValue('E' . $row, $rs->channels_name);
					$sheet->setCellValue('F' . $row, $rs->state_name);
					if (!empty($rs->reason))
					{
						$sheet->setCellValue('G' . $row, $rs->reason);
					}

					$no++;
					$row++;
				}
			}
		} //--- end WO


		if (! empty($role['WS']) or $all)
		{
			$sheet = $excel->createSheet();
			$sheet->setTitle('WS');
			$sheet->getColumnDimension('A')->setWidth(8);
			$sheet->getColumnDimension('B')->setWidth(15);
			$sheet->getColumnDimension('C')->setWidth(15);
			$sheet->getColumnDimension('D')->setWidth(15);
			$sheet->getColumnDimension('E')->setWidth(15);
			$sheet->getColumnDimension('F')->setWidth(15);
			$sheet->getColumnDimension('G')->setWidth(20);

			$sheet->setCellValue('A1', 'ลำดับ');
			$sheet->setCellValue('B1', 'วันที่');
			$sheet->setCellValue('C1', 'เลขที่');
			$sheet->setCellValue('D1', 'SAP DO');
			$sheet->setCellValue('E1', 'สถานะ');
			$sheet->setCellValue('F1', 'หมายเหตุ');

			$data = $this->document_model->getOrder('P', $fromDate, $toDate);

			if (!empty($data))
			{
				$no = 1;
				$row = 2;

				foreach ($data as $rs)
				{
					$sheet->setCellValue('A' . $row, $no);
					$sheet->setCellValue('B' . $row, thai_date($rs->date_add, FALSE, '/'));
					$sheet->setCellValue('C' . $row, $rs->code);
					$sheet->setCellValue('D' . $row, $rs->inv_code);
					$sheet->setCellValue('E' . $row, $rs->state_name);
					if (!empty($rs->reason))
					{
						$sheet->setCellValue('F' . $row, $rs->reason);
					}

					$no++;
					$row++;
				}
			}
		} //--- end WS


		if (!empty($role['WU']) or $all)
		{
			$sheet = $excel->createSheet();
			$sheet->setTitle('WU');

			//--- set Table header
			$sheet->getColumnDimension('A')->setWidth(8);
			$sheet->getColumnDimension('B')->setWidth(15);
			$sheet->getColumnDimension('C')->setWidth(15);
			$sheet->getColumnDimension('D')->setWidth(15);
			$sheet->getColumnDimension('E')->setWidth(15);
			$sheet->getColumnDimension('F')->setWidth(15);
			$sheet->getColumnDimension('G')->setWidth(20);


			$sheet->setCellValue('A1', 'ลำดับ');
			$sheet->setCellValue('B1', 'วันที่');
			$sheet->setCellValue('C1', 'เลขที่');
			$sheet->setCellValue('D1', 'SAP DO');
			$sheet->setCellValue('E1', 'สถานะ');
			$sheet->setCellValue('F1', 'หมายเหตุ');

			$data = $this->document_model->getOrder('U', $fromDate, $toDate);

			if (!empty($data))
			{
				$no = 1;
				$row = 2;

				foreach ($data as $rs)
				{
					$sheet->setCellValue('A' . $row, $no);
					$sheet->setCellValue('B' . $row, thai_date($rs->date_add, FALSE, '/'));
					$sheet->setCellValue('C' . $row, $rs->code);
					$sheet->setCellValue('D' . $row, $rs->inv_code);
					$sheet->setCellValue('E' . $row, $rs->state_name);
					if (!empty($rs->reason))
					{
						$sheet->setCellValue('F' . $row, $rs->reason);
					}

					$no++;
					$row++;
				}
			}
		} //--- end WU


		if (!empty($role['WC']) or $all)
		{
			$sheet = $excel->createSheet();
			$sheet->setTitle('WC');

			//--- set Table header
			$sheet->getColumnDimension('A')->setWidth(8);
			$sheet->getColumnDimension('B')->setWidth(15);
			$sheet->getColumnDimension('C')->setWidth(15);
			$sheet->getColumnDimension('D')->setWidth(15);
			$sheet->getColumnDimension('E')->setWidth(15);
			$sheet->getColumnDimension('F')->setWidth(15);
			$sheet->getColumnDimension('G')->setWidth(20);


			$sheet->setCellValue('A1', 'ลำดับ');
			$sheet->setCellValue('B1', 'วันที่');
			$sheet->setCellValue('C1', 'เลขที่');
			$sheet->setCellValue('D1', 'SAP DO');
			$sheet->setCellValue('E1', 'สถานะ');
			$sheet->setCellValue('F1', 'หมายเหตุ');

			$data = $this->document_model->getOrder('C', $fromDate, $toDate);

			if (!empty($data))
			{
				$no = 1;
				$row = 2;

				foreach ($data as $rs)
				{
					$sheet->setCellValue('A' . $row, $no);
					$sheet->setCellValue('B' . $row, thai_date($rs->date_add, FALSE, '/'));
					$sheet->setCellValue('C' . $row, $rs->code);
					$sheet->setCellValue('D' . $row, $rs->inv_code);
					$sheet->setCellValue('E' . $row, $rs->state_name);
					if (!empty($rs->reason))
					{
						$sheet->setCellValue('F' . $row, $rs->reason);
					}

					$no++;
					$row++;
				}
			}
		} //--- end WC


		if (!empty($role['WT']) or $all)
		{
			$sheet = $excel->createSheet();
			$sheet->setTitle('WT');

			//--- set Table header
			$sheet->getColumnDimension('A')->setWidth(8);
			$sheet->getColumnDimension('B')->setWidth(15);
			$sheet->getColumnDimension('C')->setWidth(15);
			$sheet->getColumnDimension('D')->setWidth(15);
			$sheet->getColumnDimension('E')->setWidth(15);
			$sheet->getColumnDimension('F')->setWidth(15);
			$sheet->getColumnDimension('G')->setWidth(20);

			$sheet->setCellValue('A1', 'ลำดับ');
			$sheet->setCellValue('B1', 'วันที่');
			$sheet->setCellValue('C1', 'เลขที่');
			$sheet->setCellValue('D1', 'SAP DO');
			$sheet->setCellValue('E1', 'สถานะ');
			$sheet->setCellValue('F1', 'หมายเหตุ');

			$data = $this->document_model->getOrder('N', $fromDate, $toDate);

			if (!empty($data))
			{
				$no = 1;
				$row = 2;

				foreach ($data as $rs)
				{
					$sheet->setCellValue('A' . $row, $no);
					$sheet->setCellValue('B' . $row, thai_date($rs->date_add, FALSE, '/'));
					$sheet->setCellValue('C' . $row, $rs->code);
					$sheet->setCellValue('D' . $row, $rs->inv_code);
					$sheet->setCellValue('E' . $row, $rs->state_name);
					if (!empty($rs->reason))
					{
						$sheet->setCellValue('F' . $row, $rs->reason);
					}

					$no++;
					$row++;
				}
			}
		} //--- end WT


		if (!empty($role['WQ']) or $all)
		{
			$sheet = $excel->createSheet();
			$sheet->setTitle('WQ');

			//--- set Table header
			$sheet->getColumnDimension('A')->setWidth(8);
			$sheet->getColumnDimension('B')->setWidth(15);
			$sheet->getColumnDimension('C')->setWidth(15);
			$sheet->getColumnDimension('D')->setWidth(15);
			$sheet->getColumnDimension('E')->setWidth(15);
			$sheet->getColumnDimension('F')->setWidth(15);
			$sheet->getColumnDimension('G')->setWidth(20);

			$sheet->setCellValue('A1', 'ลำดับ');
			$sheet->setCellValue('B1', 'วันที่');
			$sheet->setCellValue('C1', 'เลขที่');
			$sheet->setCellValue('D1', 'SAP TR');
			$sheet->setCellValue('E1', 'สถานะ');
			$sheet->setCellValue('F1', 'หมายเหตุ');

			$data = $this->document_model->getOrder('T', $fromDate, $toDate);

			if (!empty($data))
			{
				$no = 1;
				$row = 2;

				foreach ($data as $rs)
				{
					$sheet->setCellValue('A' . $row, $no);
					$sheet->setCellValue('B' . $row, thai_date($rs->date_add, FALSE, '/'));
					$sheet->setCellValue('C' . $row, $rs->code);
					$sheet->setCellValue('D' . $row, $rs->inv_code);
					$sheet->setCellValue('E' . $row, $rs->state_name);
					if (!empty($rs->reason))
					{
						$sheet->setCellValue('F' . $row, $rs->reason);
					}

					$no++;
					$row++;
				}
			}
		} //--- end WQ

		if (!empty($role['WV']) or $all)
		{
			$sheet = $excel->createSheet();
			$sheet->setTitle('WV');

			//--- set Table header
			$sheet->getColumnDimension('A')->setWidth(8);
			$sheet->getColumnDimension('B')->setWidth(15);
			$sheet->getColumnDimension('C')->setWidth(15);
			$sheet->getColumnDimension('D')->setWidth(15);
			$sheet->getColumnDimension('E')->setWidth(15);
			$sheet->getColumnDimension('F')->setWidth(15);
			$sheet->getColumnDimension('G')->setWidth(20);

			$sheet->setCellValue('A1', 'ลำดับ');
			$sheet->setCellValue('B1', 'วันที่');
			$sheet->setCellValue('C1', 'เลขที่');
			$sheet->setCellValue('D1', 'SAP TR');
			$sheet->setCellValue('E1', 'สถานะ');
			$sheet->setCellValue('F1', 'หมายเหตุ');

			$data = $this->document_model->getOrder('Q', $fromDate, $toDate);

			if (!empty($data))
			{
				$no = 1;
				$row = 2;

				foreach ($data as $rs)
				{
					$sheet->setCellValue('A' . $row, $no);
					$sheet->setCellValue('B' . $row, thai_date($rs->date_add, FALSE, '/'));
					$sheet->setCellValue('C' . $row, $rs->code);
					$sheet->setCellValue('D' . $row, $rs->inv_code);
					$sheet->setCellValue('E' . $row, $rs->state_name);
					if (!empty($rs->reason))
					{
						$sheet->setCellValue('F' . $row, $rs->reason);
					}

					$no++;
					$row++;
				}
			}
		} //--- end WV


		if (!empty($role['WL']) or $all)
		{
			$sheet = $excel->createSheet();
			$sheet->setTitle('WL');

			//--- set Table header
			$sheet->getColumnDimension('A')->setWidth(8);
			$sheet->getColumnDimension('B')->setWidth(15);
			$sheet->getColumnDimension('C')->setWidth(15);
			$sheet->getColumnDimension('D')->setWidth(15);
			$sheet->getColumnDimension('E')->setWidth(15);
			$sheet->getColumnDimension('F')->setWidth(15);
			$sheet->getColumnDimension('G')->setWidth(20);

			$sheet->setCellValue('A1', 'ลำดับ');
			$sheet->setCellValue('B1', 'วันที่');
			$sheet->setCellValue('C1', 'เลขที่');
			$sheet->setCellValue('D1', 'SAP TR');
			$sheet->setCellValue('E1', 'สถานะ');
			$sheet->setCellValue('F1', 'หมายเหตุ');

			$data = $this->document_model->getOrder('L', $fromDate, $toDate);

			if (!empty($data))
			{
				$no = 1;
				$row = 2;

				foreach ($data as $rs)
				{
					$sheet->setCellValue('A' . $row, $no);
					$sheet->setCellValue('B' . $row, thai_date($rs->date_add, FALSE, '/'));
					$sheet->setCellValue('C' . $row, $rs->code);
					$sheet->setCellValue('D' . $row, $rs->inv_code);
					$sheet->setCellValue('E' . $row, $rs->state_name);
					if (!empty($rs->reason))
					{
						$sheet->setCellValue('F' . $row, $rs->reason);
					}

					$no++;
					$row++;
				}
			}
		} //--- end WL


		if (!empty($role['WM']) or $all)
		{
			$sheet = $excel->createSheet();
			$sheet->setTitle('WM');

			//--- set Table header
			$sheet->getColumnDimension('A')->setWidth(8);
			$sheet->getColumnDimension('B')->setWidth(15);
			$sheet->getColumnDimension('C')->setWidth(15);
			$sheet->getColumnDimension('D')->setWidth(15);
			$sheet->getColumnDimension('E')->setWidth(15);
			$sheet->getColumnDimension('F')->setWidth(15);
			$sheet->getColumnDimension('G')->setWidth(20);

			$sheet->setCellValue('A1', 'ลำดับ');
			$sheet->setCellValue('B1', 'วันที่');
			$sheet->setCellValue('C1', 'เลขที่');
			$sheet->setCellValue('D1', 'SAP DO');
			$sheet->setCellValue('E1', 'สถานะ');
			$sheet->setCellValue('F1', 'หมายเหตุ');

			$data = $this->document_model->WM($fromDate, $toDate);

			if (!empty($data))
			{
				$no = 1;
				$row = 2;

				foreach ($data as $rs)
				{
					$sheet->setCellValue('A' . $row, $no);
					$sheet->setCellValue('B' . $row, thai_date($rs->date_add, FALSE, '/'));
					$sheet->setCellValue('C' . $row, $rs->code);
					$sheet->setCellValue('D' . $row, $rs->inv_code);
					$sheet->setCellValue('E' . $row, ($rs->status == 0 ? 'Pending' : ($rs->status == 1 ? 'Success' : ($rs->status == 2 ? 'Canceled' : 'Unknow'))));
					if (!empty($rs->reason))
					{
						$sheet->setCellValue('F' . $row, $rs->reason);
					}
					$no++;
					$row++;
				}
			}
		} //--- end WM


		if (!empty($role['WD']) or $all)
		{
			$sheet = $excel->createSheet();
			$sheet->setTitle('WD');

			//--- set Table header
			$sheet->getColumnDimension('A')->setWidth(8);
			$sheet->getColumnDimension('B')->setWidth(15);
			$sheet->getColumnDimension('C')->setWidth(15);
			$sheet->getColumnDimension('D')->setWidth(15);
			$sheet->getColumnDimension('E')->setWidth(15);
			$sheet->getColumnDimension('F')->setWidth(15);
			$sheet->getColumnDimension('G')->setWidth(20);

			$sheet->setCellValue('A1', 'ลำดับ');
			$sheet->setCellValue('B1', 'วันที่');
			$sheet->setCellValue('C1', 'เลขที่');
			$sheet->setCellValue('D1', 'SAP DO');
			$sheet->setCellValue('E1', 'สถานะ');
			$sheet->setCellValue('F1', 'หมายเหตุ');

			$data = $this->document_model->WD($fromDate, $toDate);

			if (!empty($data))
			{
				$no = 1;
				$row = 2;

				foreach ($data as $rs)
				{
					$sheet->setCellValue('A' . $row, $no);
					$sheet->setCellValue('B' . $row, thai_date($rs->date_add, FALSE, '/'));
					$sheet->setCellValue('C' . $row, $rs->code);
					$sheet->setCellValue('D' . $row, $rs->inv_code);
					$sheet->setCellValue('E' . $row, ($rs->status == 0 ? 'Pending' : ($rs->status == 1 ? 'Success' : ($rs->status == 2 ? 'Canceled' : 'Unknow'))));
					if (!empty($rs->reason))
					{
						$sheet->setCellValue('F' . $row, $rs->reason);
					}
					$no++;
					$row++;
				}
			}
		} //--- end WD


		if (!empty($role['WR']) or $all)
		{
			$sheet = $excel->createSheet();
			$sheet->setTitle('WR');

			//--- set Table header
			$sheet->getColumnDimension('A')->setWidth(8);
			$sheet->getColumnDimension('B')->setWidth(15);
			$sheet->getColumnDimension('C')->setWidth(15);
			$sheet->getColumnDimension('D')->setWidth(15);
			$sheet->getColumnDimension('E')->setWidth(15);
			$sheet->getColumnDimension('F')->setWidth(15);
			$sheet->getColumnDimension('G')->setWidth(20);

			$sheet->setCellValue('A1', 'ลำดับ');
			$sheet->setCellValue('B1', 'วันที่');
			$sheet->setCellValue('C1', 'เลขที่');
			$sheet->setCellValue('D1', 'SAP GRPO');
			$sheet->setCellValue('E1', 'สถานะ');
			$sheet->setCellValue('F1', 'หมายเหตุ');

			$data = $this->document_model->WR($fromDate, $toDate);

			if (!empty($data))
			{
				$no = 1;
				$row = 2;

				foreach ($data as $rs)
				{
					$sheet->setCellValue('A' . $row, $no);
					$sheet->setCellValue('B' . $row, thai_date($rs->date_add, FALSE, '/'));
					$sheet->setCellValue('C' . $row, $rs->code);
					$sheet->setCellValue('D' . $row, $rs->inv_code);
					$sheet->setCellValue('E' . $row, $this->statusLabel($rs->status, $rs->is_expire));
					if ($rs->status == 2)
					{
						$sheet->setCellValue('F' . $row, $rs->reason);
					}

					$no++;
					$row++;
				}
			}
		} //--- end WR


		if (!empty($role['WW']) or $all)
		{
			$sheet = $excel->createSheet();
			$sheet->setTitle('WW');

			//--- set Table header
			$sheet->getColumnDimension('A')->setWidth(8);
			$sheet->getColumnDimension('B')->setWidth(15);
			$sheet->getColumnDimension('C')->setWidth(15);
			$sheet->getColumnDimension('D')->setWidth(15);
			$sheet->getColumnDimension('E')->setWidth(15);
			$sheet->getColumnDimension('F')->setWidth(15);
			$sheet->getColumnDimension('G')->setWidth(20);

			$sheet->setCellValue('A1', 'ลำดับ');
			$sheet->setCellValue('B1', 'วันที่');
			$sheet->setCellValue('C1', 'เลขที่');
			$sheet->setCellValue('D1', 'SAP TR');
			$sheet->setCellValue('E1', 'สถานะ');
			$sheet->setCellValue('F1', 'หมายเหตุ');

			$data = $this->document_model->WW($fromDate, $toDate);

			if (!empty($data))
			{
				$no = 1;
				$row = 2;

				foreach ($data as $rs)
				{
					$sheet->setCellValue('A' . $row, $no);
					$sheet->setCellValue('B' . $row, thai_date($rs->date_add, FALSE, '/'));
					$sheet->setCellValue('C' . $row, $rs->code);
					$sheet->setCellValue('D' . $row, $rs->inv_code);
					$sheet->setCellValue('E' . $row, $this->statusLabel($rs->status, $rs->is_expire));
					if (!empty($rs->reason))
					{
						$sheet->setCellValue('F' . $row, $rs->reason);
					}
					$no++;
					$row++;
				}
			}
		} //--- end WW


		if (!empty($role['MV']) or $all)
		{
			$sheet = $excel->createSheet();
			$sheet->setTitle('MV');

			//--- set Table header
			$sheet->getColumnDimension('A')->setWidth(8);
			$sheet->getColumnDimension('B')->setWidth(15);
			$sheet->getColumnDimension('C')->setWidth(15);
			$sheet->getColumnDimension('D')->setWidth(15);
			$sheet->getColumnDimension('E')->setWidth(15);
			$sheet->getColumnDimension('F')->setWidth(15);
			$sheet->getColumnDimension('G')->setWidth(20);

			$sheet->setCellValue('A1', 'ลำดับ');
			$sheet->setCellValue('B1', 'วันที่');
			$sheet->setCellValue('C1', 'เลขที่');
			$sheet->setCellValue('D1', 'SAP TR');
			$sheet->setCellValue('E1', 'สถานะ');
			$sheet->setCellValue('F1', 'หมายเหตุ');

			$data = $this->document_model->MV($fromDate, $toDate);

			if (!empty($data))
			{
				$no = 1;
				$row = 2;

				foreach ($data as $rs)
				{
					$sheet->setCellValue('A' . $row, $no);
					$sheet->setCellValue('B' . $row, thai_date($rs->date_add, FALSE, '/'));
					$sheet->setCellValue('C' . $row, $rs->code);
					$sheet->setCellValue('D' . $row, $rs->inv_code);
					$sheet->setCellValue('E' . $row, $this->statusLabel($rs->status, $rs->is_expire));
					if (!empty($rs->reason))
					{
						$sheet->setCellValue('F' . $row, $rs->reason);
					}

					$no++;
					$row++;
				}
			}
		} //--- end MV


		if (!empty($role['WG']) or $all)
		{
			$sheet = $excel->createSheet();
			$sheet->setTitle('WG');

			//--- set Table header
			$sheet->getColumnDimension('A')->setWidth(8);
			$sheet->getColumnDimension('B')->setWidth(15);
			$sheet->getColumnDimension('C')->setWidth(15);
			$sheet->getColumnDimension('D')->setWidth(15);
			$sheet->getColumnDimension('E')->setWidth(15);
			$sheet->getColumnDimension('F')->setWidth(15);
			$sheet->getColumnDimension('G')->setWidth(20);

			$sheet->setCellValue('A1', 'ลำดับ');
			$sheet->setCellValue('B1', 'วันที่');
			$sheet->setCellValue('C1', 'เลขที่');
			$sheet->setCellValue('D1', 'Base Ref');
			$sheet->setCellValue('E1', 'SAP Goods Issue');
			$sheet->setCellValue('F1', 'สถานะ');
			$sheet->setCellValue('G1', 'หมายเหตุ');

			$data = $this->document_model->WG($fromDate, $toDate);

			if (!empty($data))
			{
				$no = 1;
				$row = 2;

				foreach ($data as $rs)
				{
					$sheet->setCellValue('A' . $row, $no);
					$sheet->setCellValue('B' . $row, thai_date($rs->date_add, FALSE, '/'));
					$sheet->setCellValue('C' . $row, $rs->code);
					$sheet->setCellValue('D' . $row, $rs->reference);
					$sheet->setCellValue('E' . $row, $rs->inv_code);
					$sheet->setCellValue('F' . $row, ($rs->status == 0 ? 'Pending' : ($rs->status == 1 ? 'Success' : ($rs->status == 2 ? 'Canceled' : 'WMS'))));
					if (!empty($rs->reason))
					{
						$sheet->setCellValue('G' . $row, $rs->reason);
					}
					$no++;
					$row++;
				}
			}
		} //--- end WG


		if (!empty($role['RT']) or $all)
		{
			$sheet = $excel->createSheet();
			$sheet->setTitle('RT');

			//--- set Table header
			$sheet->getColumnDimension('A')->setWidth(8);
			$sheet->getColumnDimension('B')->setWidth(15);
			$sheet->getColumnDimension('C')->setWidth(15);
			$sheet->getColumnDimension('D')->setWidth(15);
			$sheet->getColumnDimension('E')->setWidth(15);
			$sheet->getColumnDimension('F')->setWidth(15);
			$sheet->getColumnDimension('G')->setWidth(20);

			$sheet->setCellValue('A1', 'ลำดับ');
			$sheet->setCellValue('B1', 'วันที่');
			$sheet->setCellValue('C1', 'เลขที่');
			$sheet->setCellValue('D1', 'Base Ref');
			$sheet->setCellValue('E1', 'SAP Goods Issue');
			$sheet->setCellValue('F1', 'สถานะ');
			$sheet->setCellValue('G1', 'หมายเหตุ');

			$data = $this->document_model->RT($fromDate, $toDate);

			if (!empty($data))
			{
				$no = 1;
				$row = 2;

				foreach ($data as $rs)
				{
					$sheet->setCellValue('A' . $row, $no);
					$sheet->setCellValue('B' . $row, thai_date($rs->date_add, FALSE, '/'));
					$sheet->setCellValue('C' . $row, $rs->code);
					$sheet->setCellValue('D' . $row, $rs->reference);
					$sheet->setCellValue('E' . $row, $rs->inv_code);
					$sheet->setCellValue('F' . $row, $this->statusLabel($rs->status, $rs->is_expire));
					if (!empty($rs->reason))
					{
						$sheet->setCellValue('G' . $row, $rs->reason);
					}
					$no++;
					$row++;
				}
			}
		} //--- end RT


		if (!empty($role['RN']) or $all)
		{
			$sheet = $excel->createSheet();
			$sheet->setTitle('RN');

			//--- set Table header
			$sheet->getColumnDimension('A')->setWidth(8);
			$sheet->getColumnDimension('B')->setWidth(15);
			$sheet->getColumnDimension('C')->setWidth(15);
			$sheet->getColumnDimension('D')->setWidth(15);
			$sheet->getColumnDimension('E')->setWidth(15);
			$sheet->getColumnDimension('F')->setWidth(15);
			$sheet->getColumnDimension('G')->setWidth(20);

			$sheet->setCellValue('A1', 'ลำดับ');
			$sheet->setCellValue('B1', 'วันที่');
			$sheet->setCellValue('C1', 'เลขที่');
			$sheet->setCellValue('D1', 'Base Ref');
			$sheet->setCellValue('E1', 'SAP TR');
			$sheet->setCellValue('F1', 'สถานะ');
			$sheet->setCellValue('G1', 'หมายเหตุ');

			$data = $this->document_model->RN($fromDate, $toDate);

			if (!empty($data))
			{
				$no = 1;
				$row = 2;

				foreach ($data as $rs)
				{
					$sheet->setCellValue('A' . $row, $no);
					$sheet->setCellValue('B' . $row, thai_date($rs->date_add, FALSE, '/'));
					$sheet->setCellValue('C' . $row, $rs->code);
					$sheet->setCellValue('D' . $row, $rs->reference);
					$sheet->setCellValue('E' . $row, $rs->inv_code);
					$sheet->setCellValue('F' . $row, $this->statusLabel($rs->status, $rs->is_expire));
					if (!empty($rs->reason))
					{
						$sheet->setCellValue('G' . $row, $rs->reason);
					}
					$no++;
					$row++;
				}
			}
		} //--- end RN


		if (!empty($role['SM']) or $all)
		{
			$sheet = $excel->createSheet();
			$sheet->setTitle('SM');

			//--- set Table header
			$sheet->getColumnDimension('A')->setWidth(8);
			$sheet->getColumnDimension('B')->setWidth(15);
			$sheet->getColumnDimension('C')->setWidth(15);
			$sheet->getColumnDimension('D')->setWidth(15);
			$sheet->getColumnDimension('E')->setWidth(15);
			$sheet->getColumnDimension('F')->setWidth(15);
			$sheet->getColumnDimension('G')->setWidth(20);

			$sheet->setCellValue('A1', 'ลำดับ');
			$sheet->setCellValue('B1', 'วันที่');
			$sheet->setCellValue('C1', 'เลขที่');
			$sheet->setCellValue('D1', 'Base Ref');
			$sheet->setCellValue('E1', 'SAP SM');
			$sheet->setCellValue('F1', 'สถานะ');
			$sheet->setCellValue('G1', 'หมายเหตุ');

			$data = $this->document_model->SM($fromDate, $toDate);

			if (!empty($data))
			{
				$no = 1;
				$row = 2;

				foreach ($data as $rs)
				{
					$sheet->setCellValue('A' . $row, $no);
					$sheet->setCellValue('B' . $row, thai_date($rs->date_add, FALSE, '/'));
					$sheet->setCellValue('C' . $row, $rs->code);
					$sheet->setCellValue('D' . $row, $rs->reference);
					$sheet->setCellValue('E' . $row, $rs->inv_code);
					$sheet->setCellValue('F' . $row, $this->statusLabel($rs->status, $rs->is_expire));
					if (!empty($rs->reason))
					{
						$sheet->setCellValue('G' . $row, $rs->reason);
					}
					$no++;
					$row++;
				}
			}
		} //--- end SM


		if (!empty($role['CN']) or $all)
		{
			$sheet = $excel->createSheet();
			$sheet->setTitle('CN');

			//--- set Table header
			$sheet->getColumnDimension('A')->setWidth(8);
			$sheet->getColumnDimension('B')->setWidth(15);
			$sheet->getColumnDimension('C')->setWidth(15);
			$sheet->getColumnDimension('D')->setWidth(15);
			$sheet->getColumnDimension('E')->setWidth(15);
			$sheet->getColumnDimension('F')->setWidth(15);
			$sheet->getColumnDimension('G')->setWidth(20);

			$sheet->setCellValue('A1', 'ลำดับ');
			$sheet->setCellValue('B1', 'วันที่');
			$sheet->setCellValue('C1', 'เลขที่');
			$sheet->setCellValue('D1', 'Base Ref');
			$sheet->setCellValue('E1', 'SAP SM');
			$sheet->setCellValue('F1', 'สถานะ');
			$sheet->setCellValue('G1', 'หมายเหตุ');

			$data = $this->document_model->CN($fromDate, $toDate);

			if (!empty($data))
			{
				$no = 1;
				$row = 2;

				foreach ($data as $rs)
				{
					$sheet->setCellValue('A' . $row, $no);
					$sheet->setCellValue('B' . $row, thai_date($rs->date_add, FALSE, '/'));
					$sheet->setCellValue('C' . $row, $rs->code);
					$sheet->setCellValue('D' . $row, $rs->reference);
					$sheet->setCellValue('E' . $row, $rs->inv_code);
					$sheet->setCellValue('F' . $row, ($rs->status == 0 ? 'Pending' : ($rs->status == 1 ? 'Success' : ($rs->status == 2 ? 'Canceled' : 'WMS'))));
					if (!empty($rs->reason))
					{
						$sheet->setCellValue('G' . $row, $rs->reason);
					}
					$no++;
					$row++;
				}
			}
		} //--- end CN


		if (!empty($role['WA']) or $all)
		{
			$sheet = $excel->createSheet();
			$sheet->setTitle('WA');

			//--- set Table header
			$sheet->getColumnDimension('A')->setWidth(8);
			$sheet->getColumnDimension('B')->setWidth(15);
			$sheet->getColumnDimension('C')->setWidth(15);
			$sheet->getColumnDimension('D')->setWidth(15);
			$sheet->getColumnDimension('E')->setWidth(15);
			$sheet->getColumnDimension('F')->setWidth(15);
			$sheet->getColumnDimension('G')->setWidth(20);

			$sheet->setCellValue('A1', 'ลำดับ');
			$sheet->setCellValue('B1', 'วันที่');
			$sheet->setCellValue('C1', 'เลขที่');
			$sheet->setCellValue('D1', 'SAP Goods Issue');
			$sheet->setCellValue('E1', 'SAP Goods Receipt');
			$sheet->setCellValue('F1', 'สถานะ');
			$sheet->setCellValue('G1', 'หมายเหตุ');

			$data = $this->document_model->WA($fromDate, $toDate);

			if (!empty($data))
			{
				$no = 1;
				$row = 2;

				foreach ($data as $rs)
				{
					$sheet->setCellValue('A' . $row, $no);
					$sheet->setCellValue('B' . $row, thai_date($rs->date_add, FALSE, '/'));
					$sheet->setCellValue('C' . $row, $rs->code);
					$sheet->setCellValue('D' . $row, $rs->issue_code);
					$sheet->setCellValue('E' . $row, $rs->receive_code);
					$sheet->setCellValue('F' . $row, ($rs->status == 0 ? 'Pending' : ($rs->status == 1 ? 'Success' : ($rs->status == 2 ? 'Canceled' : 'Unknow'))));
					if (!empty($rs->reason))
					{
						$sheet->setCellValue('G' . $row, $rs->reason);
					}
					$no++;
					$row++;
				}
			}
		} //--- end WA

		if (!empty($role['AC']) or $all)
		{
			$sheet = $excel->createSheet();
			$sheet->setTitle('AC');

			//--- set Table header
			$sheet->getColumnDimension('A')->setWidth(8);
			$sheet->getColumnDimension('B')->setWidth(15);
			$sheet->getColumnDimension('C')->setWidth(15);
			$sheet->getColumnDimension('D')->setWidth(15);
			$sheet->getColumnDimension('E')->setWidth(15);
			$sheet->getColumnDimension('F')->setWidth(15);
			$sheet->getColumnDimension('G')->setWidth(15);
			$sheet->getColumnDimension('H')->setWidth(15);

			$sheet->setCellValue('A1', 'ลำดับ');
			$sheet->setCellValue('B1', 'วันที่');
			$sheet->setCellValue('C1', 'เลขที่');
			$sheet->setCellValue('D1', 'Base ref');
			$sheet->setCellValue('E1', 'SAP Goods Issue');
			$sheet->setCellValue('F1', 'SAP Goods Receipt');
			$sheet->setCellValue('G1', 'สถานะ');
			$sheet->setCellValue('H1', 'หมายเหตุ');

			$data = $this->document_model->AC($fromDate, $toDate);

			if (!empty($data))
			{
				$no = 1;
				$row = 2;

				foreach ($data as $rs)
				{
					$sheet->setCellValue('A' . $row, $no);
					$sheet->setCellValue('B' . $row, thai_date($rs->date_add, FALSE, '/'));
					$sheet->setCellValue('C' . $row, $rs->code);
					$sheet->setCellValue('D' . $row, $rs->reference);
					$sheet->setCellValue('E' . $row, $rs->issue_code);
					$sheet->setCellValue('F' . $row, $rs->receive_code);
					$sheet->setCellValue('G' . $row, ($rs->status == 0 ? 'Pending' : ($rs->status == 1 ? 'Success' : ($rs->status == 2 ? 'Canceled' : 'Unknow'))));
					if (!empty($rs->reason))
					{
						$sheet->setCellValue('H' . $row, $rs->reason);
					}
					$no++;
					$row++;
				}
			}
		} //--- end AC


		setToken($token);
		$file_name = "รายงานทะเบียนคุมเอกสารแยกตามประเภท " . date('Ymd') . ".xlsx";
		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); /// form excel 2007 XLSX
		header('Content-Disposition: attachment;filename="' . $file_name . '"');
		$writer = new Xlsx($excel);
		$writer->save('php://output');
	}


	public function statusLabel($status, $is_expire = 0)
	{
		$label = "Unknow";

		if ($is_expire == 1)
		{
			$label = "Expired";
		}
		else
		{
			switch ($status)
			{
				case -1:
					$label = "Draft";
					break;
				case 0:
					$label = "Pending";
					break;
				case 1:
					$label = "Success";
					break;
				case 2:
					$label = "Canceled";
					break;
				case 3:
					$label = "WMS";
					break;
				case 4:
					$label = "Acception";
					break;
			}
		}

		return $label;
	}
}
