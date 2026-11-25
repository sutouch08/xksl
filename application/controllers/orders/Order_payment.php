<?php
defined('BASEPATH') OR exit('No direct script access allowed');
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class Order_payment extends PS_Controller
{
  public $menu_code = 'ACPMCF';
	public $menu_group_code = 'AC';
  public $menu_sub_group_code = '';
	public $title = 'ตรวจสอบยอดชำระเงิน';
  public $filter;
	public $error;

  public function __construct()
  {
    parent::__construct();
    $this->home = base_url().'orders/order_payment';
    $this->load->model('orders/order_payment_model');
    $this->load->model('masters/bank_model');
    $this->load->helper('bank');
    $this->load->helper('order');
    $this->load->helper('channels');
  }


  public function index()
  {
    $filter = array(
      'code'  => get_filter('code', 'code', ''),
      'customer' => get_filter('customer', 'customer', ''),
      'account' => get_filter('account', 'account', 'all'),
      'user'  => get_filter('user', 'user', ''),
      'channels' => get_filter('channels', 'channels', 'all'),
      'from_date' => get_filter('from_date', 'from_date', ''),
      'to_date'  => get_filter('to_date', 'to_date', ''),
      'valid' => get_filter('valid', 'valid', '0'),
      'is_pre_order' => get_filter('is_pre_order', 'is_pre_order', 'all')
    );

    if($this->input->post('search'))
    {
      redirect($this->home);
    }
    else
    {
      $perpage = get_rows();      
      $segment  = 4;
      $rows = $this->order_payment_model->count_rows($filter);
      $init = pagination_config($this->home.'/index/', $rows, $perpage, $segment);
      $orders = $this->order_payment_model->get_list($filter, $perpage, $this->uri->segment($segment));
      $filter['orders'] = $orders;
      $this->pagination->initialize($init);
      $this->load->view('orders/payment/order_payment_list', $filter);
    }
  }


  public function get_payment_detail()
  {
    $sc = TRUE;
    $id = $this->input->post('id');
    $detail = $this->order_payment_model->get_detail($id);
    if(!empty($detail))
    {
      $img = payment_image_url($detail->order_code);
      $bank   = $this->bank_model->get_account_detail($detail->id_account);
      $ds  = array(
        'id' => $detail->id,
        'orderAmount' => number($detail->order_amount,2),
        'payAmount' => number($detail->pay_amount,2),
        'payDate' => thai_date($detail->pay_date, TRUE, '/'),
        'bankName' => $bank->bank_name,
        'branch' => $bank->branch,
        'accNo' => $bank->acc_no,
        'accName' => $bank->acc_name,
        'date_add' => thai_date($detail->date_upd, TRUE, '/'),
        'imageUrl' => $img
      );

      if($detail->valid == 0)
      {
        $ds['valid'] = 'no';
      }
    }
    else
    {
      $sc = FALSE;
    }

    echo $sc === TRUE ? json_encode($ds) : 'fail';
  }


  public function confirm_payment()
  {
    $sc = TRUE;

    if($this->input->post('id'))
    {
      $this->load->model('orders/orders_model');
      $this->load->model('orders/order_state_model');
      $id = $this->input->post('id');
      $detail = $this->order_payment_model->get_detail($id);
			$order = $this->orders_model->get($detail->order_code);

      $arr = array(
        'order_code' => $detail->order_code,
        'state' => 3,
        'update_user' => $this->_user->uname
      );

      $this->db->trans_begin();

      $this->order_payment_model->valid_payment($id);

      $this->orders_model->paid($detail->order_code, TRUE);

			if($order->state < 3)
			{
	      $this->orders_model->change_state($detail->order_code, 3);

	      $this->order_state_model->add_state($arr);
			}

			if($sc === TRUE)
			{
				$this->db->trans_commit();
			}
			else
			{
				$this->db->trans_rollback();
			}
    }
    else
    {
      $sc = FALSE;
      $this->error = 'ไม่พบรายการชำระเงิน';
    }

    echo $sc === TRUE ? 'success' : $this->error;
  }



  public function un_confirm_payment()
  {
    $sc = TRUE;

    if($this->input->post('id'))
    {
      $this->load->model('orders/orders_model');
      $this->load->model('orders/order_state_model');
      $id = $this->input->post('id');
      $detail = $this->order_payment_model->get_detail($id);
			$order = $this->orders_model->get($detail->order_code);

      $arr = array(
        'order_code' => $detail->order_code,
        'state' => 2,
        'update_user' => get_cookie('uname')
      );

      $this->db->trans_start();

      $this->order_payment_model->un_valid_payment($id);

      $this->orders_model->paid($detail->order_code, FALSE);

			if($order->state != 8 && $order->state != 9)
			{
	      $this->orders_model->change_state($detail->order_code, 2);

	      $this->order_state_model->add_state($arr);
			}

	    $this->db->trans_complete();

	    if($this->db->trans_status() === FALSE)
	    {
	      $sc = FALSE;
	      $message = $this->db->error();
	    }
    }
    else
    {
      $sc = FALSE;
      $message = 'ไม่พบรายการชำระเงิน';
    }

    echo $sc === TRUE ? 'success' : $message;
  }


  public function remove_payment()
  {
    $sc = TRUE;
    if($this->input->post('id'))
    {
      $this->load->model('orders/orders_model');
      $this->load->model('orders/order_state_model');
      $id = $this->input->post('id');
      $detail = $this->order_payment_model->get_detail($id);

      if(!empty($detail))
      {
				$order = $this->orders_model->get($detail->order_code);

        $this->db->trans_start();

        $this->orders_model->paid($detail->order_code, FALSE);

				if($order->state != 8 && $order->state != 9)
				{
	        $this->orders_model->change_state($detail->order_code, 1);

	        $arr = array(
	          'order_code' => $detail->order_code,
	          'state' => 1,
	          'update_user' => get_cookie('uname')
	        );

	        $this->order_state_model->add_state($arr);
				}

        $this->order_payment_model->delete($id);

        $this->db->trans_complete();

        if($this->db->trans_status() === FALSE)
        {
          $sc = FALSE;
          $message = $this->db->error();
        }
      }
      else
      {
        $sc = FALSE;
        $message = 'ไม่พบรายการชำระเงิน';
      }
    }
    else
    {
      $sc = FALSE;
      $message = 'ไม่พบตัวแปร id กรุณา reload หน้าเว็บแล้วลองใหม่';
    }

    echo $sc === TRUE ? 'success' : $message;
  }

  private function novalue($value = NULL)
  {
    $re = ($value == NULL OR $value == "") ? "ไม่ระบุ" : $value;
    return $re;
  }


  public function export_filter()
  {
    $this->load->model('masters/channels_model');

    $filter = array(
      'code'  => get_filter('code', 'code', ''),
      'customer' => get_filter('customer', 'customer', ''),
      'account' => get_filter('account', 'account', 'all'),
      'user'  => get_filter('user', 'user', ''),
      'channels' => get_filter('channels', 'channels', 'all'),
      'from_date' => get_filter('from_date', 'from_date', ''),
      'to_date'  => get_filter('to_date', 'to_date', ''),
      'valid' => get_filter('valid', 'valid', '0'),
      'is_pre_order' => get_filter('is_pre_order', 'is_pre_order', 'all')
    );

    $token = $this->input->post('token');

    $data = $this->order_payment_model->get_data($filter);

    $excel = new Spreadsheet();
    $sheet = $excel->getActiveSheet();

    $excel->setActiveSheetIndex(0);
    $sheet->setTitle('ตรวจสอบยอดชำระเงิน');

    $sheet->setCellValue('A1', "รายกงานตรวจสอบยอดชำระเงิน @".date('d/m/Y'));
    $sheet->mergeCells('A1:L1');
    $row = 2;
    $date = empty($filter['from_date'] OR $filter['to_date']) ? 'ไม่ระบุ' : $filter['from_date'].' ถึง '.$filter['to_date'];

    $bank = $filter['account'] == 'all' ? NULL : $this->bank_model->get_detail($filter['account']);
    $acc = $filter['account'] == 'all' ? 'ทั้งหมด' : (empty($bank) ? "" : $bank->account_no);

    $channels = $filter['channels'] == 'all' ? 'ทั้งหมด' : $this->channels_model->get_name($filter['channels']);

    $sheet->setCellValue('A'.$row, 'Filter');
    $sheet->setCellValue('B'.$row, $this->novalue($filter['code']));
    $sheet->setCellValue('C'.$row, $filter['is_pre_order'] == '1' ? 'Y' : ($filter['is_pre_order'] == '0' ? 'N' : 'ทั้งหมด'));
    $sheet->setCellValue('D'.$row, $channels);
    $sheet->setCellValue('E'.$row, $this->novalue($filter['customer']));
    $sheet->setCellValue('F'.$row, $this->novalue($filter['user']));
    $sheet->setCellValue('G'.$row, $date);
    $sheet->setCellValue('I'.$row, $acc);
    $sheet->setCellValue('J'.$row, $filter['valid'] == 'all' ? 'ทั้งหมด' : ($filter['valid'] == '1' ? 'ยืนยันแล้ว' : 'รอยืนยัน'));

		$row = 3;

    $sheet->setCellValue('A'.$row, '#');
    $sheet->setCellValue('B'.$row, 'เลขที่เอกสาร');
    $sheet->setCellValue('C'.$row, 'Pre order');
    $sheet->setCellValue('D'.$row, 'ช่องทางขาย');
    $sheet->setCellValue('E'.$row, 'ลูกค้า');
    $sheet->setCellValue('F'.$row, 'พนักงาน');
    $sheet->setCellValue('G'.$row, 'วัน-เวลา');
    $sheet->setCellValue('H'.$row, 'ยอดเงิน');
    $sheet->setCellValue('I'.$row, 'เลขที่บัญชี');
    $sheet->setCellValue('J'.$row, 'สถานะ');
    $sheet->setCellValue('K'.$row, 'ยืนยันโดย');
    $sheet->setCellValue('L'.$row, 'วันที่ยืนยัน');

    $sheet->getColumnDimension('B')->setWidth(20);
    $sheet->getColumnDimension('C')->setWidth(20);
    $sheet->getColumnDimension('D')->setWidth(20);
    $sheet->getColumnDimension('E')->setWidth(20);
    $sheet->getColumnDimension('F')->setWidth(30);
    $sheet->getColumnDimension('G')->setWidth(30);
    $sheet->getColumnDimension('H')->setWidth(15);
    $sheet->getColumnDimension('I')->setWidth(15);
    $sheet->getColumnDimension('J')->setWidth(15);
    $sheet->getColumnDimension('K')->setWidth(15);
    $sheet->getColumnDimension('L')->setWidth(15);

		$row++;


    if( ! empty($data))
    {
      $no = 1;

      foreach($data as $rs)
      {
        $customer_name = ( ! empty($rs->customer_ref)) ? $rs->customer_ref : $rs->customer_name;

        $sheet->setCellValue('A'.$row, $no);
        $sheet->setCellValue('B'.$row, $rs->order_code);
        $sheet->setCellValue('C'.$row, ($rs->is_pre_order ? 'Y' : 'N'));
        $sheet->setCellValue('D'.$row, $rs->channels);
        $sheet->setCellValue('E'.$row, $customer_name);
        $sheet->setCellValue('F'.$row, $rs->user);
        $sheet->setCellValue('G'.$row, thai_date($rs->pay_date, TRUE));
        $sheet->setCellValue('H'.$row, $rs->pay_amount);
        $sheet->setCellValue('I'.$row, $rs->acc_no);
        $sheet->setCellValue('J'.$row, ($rs->valid == 1 ? 'ยืนยันแล้ว' : 'รอยืนยัน'));
        $sheet->setCellValue('K'.$row, $rs->validate_by);
        $sheet->setCellValue('L'.$row, thai_date($rs->date_upd, TRUE));

        $no++;
        $row++;
      }

      $sheet->getStyle("H4:H{$row}")->getNumberFormat()->setFormatCode('#,##0.00');

    }
		else
		{
			$sheet->setCellValue('A'.$row, "ไม่พบข้อมูลตามเงื่อนไขที่ระบุ");
		}

    setToken($token);
    $file_name = "รายงานตรวจสอบยอดชำระเงิน_".date('dmY').".xlsx";
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); /// form excel 2007 XLSX
    header('Content-Disposition: attachment;filename="'.$file_name.'"');
    $writer = new Xlsx($excel);
    $writer->save('php://output');
  }


  public function clear_filter()
  {
    $filter = array('code', 'account', 'user', 'channels','from_date', 'to_date', 'customer', 'valid', 'is_pre_order');
    clear_filter($filter);
  }
} //--- end class

?>
