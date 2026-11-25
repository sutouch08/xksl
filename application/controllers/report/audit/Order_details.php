<?php
defined('BASEPATH') OR exit('No direct script access allowed');
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class Order_details extends PS_Controller
{
  public $menu_code = 'RAODDT';
	public $menu_group_code = 'RE';
  public $menu_sub_group_code = 'REAUDIT';
	public $title = 'รายงานสถานะเอกสารขาออก แสดงรายละเอียด';
  
  public function __construct()
  {
    parent::__construct();
    $this->home = base_url().'report/audit/order_details';
    $this->load->model('report/audit/order_details_model');
    $this->load->model('masters/products_model');
    $this->load->model('orders/orders_model');
    $this->load->model('masters/channels_model');
    $this->load->model('masters/payment_methods_model');
    $this->load->model('masters/warehouse_model');
    $this->load->model('masters/customers_model');
  }

  public function index()
  {
    $ds = array(
      'channels_list' => $this->channels_model->get_data(),
      'payment_list' => $this->payment_methods_model->get_data(),
      'warehouse_list' => $this->warehouse_model->get_sell_warehouse_list()
    );

    $this->load->view('report/audit/report_order_details', $ds);
  }

  private function channels_array()
  {
    $ds = array();

    $rs = $this->db->get('channels');
    if($rs->num_rows() > 0)
    {
      foreach($rs->result() as $rd)
      {
        $ds[$rd->code] = $rd->name;
      }
    }

    return $ds;
  }

  private function payment_array()
  {
    $ds = array();
    $rs = $this->db->get('payment_method');

    if($rs->num_rows() > 0)
    {
      foreach($rs->result() as $rd)
      {
        $ds[$rd->code] = $rd->name;
      }
    }

    return $ds;
  }

  private function warehouse_array()
  {
    $ds = array();
    $rs = $this->db->get('warehouse');

    if($rs->num_rows() > 0)
    {
      foreach($rs->result() as $rd)
      {
        $ds[$rd->code] = $rd->name;
      }
    }

    return $ds;
  }

  private function state_array()
  {
    $ds = array(
      '1' => 'รอดำเนินการ',
      '2' => 'รอชำระเงิน',
      '3' => 'รอจัดสินค้า',
      '4' => 'กำลังจัดสินค้า',
      '5' => 'รอตรวจสินค้า',
      '6' => 'กำลังตรวจสินค้า',
      '7' => 'รอการจัดส่ง',
      '8' => 'จัดส่งแล้ว',
      '9' => 'ยกเลิก'
    );

    return $ds;
  }

  private function cancel_reason($code)
  {
    $rs = $this->db->where('code', $code)->order_by('id', 'DESC')->limit(1)->get('order_cancle_reason');

    if($rs->num_rows() == 1)
    {
      return $rs->row()->reason;
    }

    return NULL;
  }


  public function get_report()
  {
    ini_set('memory_limit','2048M'); // This also needs to be increased in some cases. Can be changed to a higher value as per need)
    ini_set('sqlsrv.ClientBufferMaxKBSize','2097152'); // Setting to 2048M
    ini_set('sqlsrv.client_buffer_max_kb_size','2097152'); // Setting to 512M - for pdo_sqlsrv

    $sc = TRUE;

    if($this->input->get("json"))
    {
      $json = json_decode($this->input->get("json") ?? '{}');

      if( ! empty($json))
      {
        $filter = array(
          "from_date" => $json->fromDate,
          "to_date" => $json->toDate,
          "all_role" => $json->allRole,
          "role" => $json->role,
          "is_expired" => $json->isExpired,
          "is_preorder" => $json->isPreOrder,
          "all_state" => $json->allState,
          "state" => $json->state,
          "all_channels" => $json->allChannels,
          "channels" => $json->channels,
          "all_payment" => $json->allPayments,
          "payment" => $json->payment,
          "all_warehouse" => $json->allWarehouse,
          "warehouse" => $json->warehouse
        );

        $limit = 1000;
        $rows = $this->order_details_model->count_rows($filter);
        $details = $this->order_details_model->get_data($filter, $limit);
        $ds = array();

        if( ! empty($details))
        {
          $state = $this->state_array();
          $wh = $this->warehouse_array();
          $channels = $this->channels_array();
          $payment = $this->payment_array();
          $cust = array();

          $no = 1;

          foreach($details as $rs)
          {
            $rs->no = number($no);
            $rs->expired = $rs->is_expired == 1 ? 'Y' : 'N';
            $rs->is_preorder = $rs->is_pre_order == 1 ? 'Y' : 'N';
            $rs->date_add = thai_date($rs->date_add, FALSE, '/');
            $rs->date_upd = thai_date($rs->date_upd, FALSE, '/');
            $rs->total_amount = number($this->order_details_model->get_doc_total($rs->code), 2);
            $rs->channels_name = empty($channels[$rs->channels_code]) ? NULL : $channels[$rs->channels_code];
            $rs->payment_name = empty($payment[$rs->payment_code]) ? NULL : $payment[$rs->payment_code];
            $rs->warehouse_name = empty($wh[$rs->warehouse_code]) ? NULL : $wh[$rs->warehouse_code];
            $rs->state_name = $state[$rs->state];
            $rs->uname = $rs->user;
            $rs->emp_name = $this->user_model->get_name($rs->user);
            $rs->cancel_reason = $rs->state == 9 ? $this->cancel_reason($rs->code) : NULL;

            if(empty($cust[$rs->customer_code]))
            {
              $customer_name = $this->customers_model->get_name($rs->customer_code);
              $cust[$rs->customer_code] = $customer_name;
            }

            $rs->customer_name = $cust[$rs->customer_code];
            $no++;
          }
        }

        $ds = array(
          "limit" => $limit,
          "rows" => $rows,
          "details" => $details
        );
      }
      else
      {
        $sc = FALSE;
        $this->error = "Invalid request data format";
      }
    }

    echo $sc === TRUE ? json_encode($ds) : $this->error;
  }



  public function do_export()
  {
    ini_set('memory_limit','2048M'); // This also needs to be increased in some cases. Can be changed to a higher value as per need)
    ini_set('sqlsrv.ClientBufferMaxKBSize','2097152'); // Setting to 2048M
    ini_set('sqlsrv.client_buffer_max_kb_size','2097152'); // Setting to 512M - for pdo_sqlsrv
    
    $token = $this->input->post('token');

    $excel = new Spreadsheet();
    $sheet = $excel->getActiveSheet();
    $sheet->setTitle('Order details Report');

    if($this->input->post('fromDate'))
    {
      $allRole = $this->input->post('allRole');
      $allWarehouse = $this->input->post('allWarehouse');
      $allChannels = $this->input->post('allChannels');
      $allPayments = $this->input->post('allPayments');
      $allState = $this->input->post('allState');
      $is_expired = $this->input->post('is_expired');
      $is_preorder = $this->input->post('is_preorder');
      $fromDate = $this->input->post('fromDate');
      $toDate = $this->input->post('toDate');
      $role = $this->input->post('role');
      $channels = $this->input->post('channels');
      $payments = $this->input->post('payments');
      $warehouse = $this->input->post('warehouse');
      $state = $this->input->post('state');

      $filter = array(
        "from_date" => $fromDate,
        "to_date" => $toDate,
        "all_role" => $allRole,
        "role" => $role,
        "is_expired" => $is_expired,
        "is_preorder" => $is_preorder,
        "all_state" => $allState,
        "state" => $state,
        "all_channels" => $allChannels,
        "channels" => $channels,
        "all_payment" => $allPayments,
        "payment" => $payments,
        "all_warehouse" => $allWarehouse,
        "warehouse" => $warehouse
      );

      $details = $this->order_details_model->get_data($filter);

      //---  Report title
      $report_title = "รายงานสถานะเอกสารแสดงรายละเอียด";
      $role_title = "เอกสาร : ". ($allRole == 1 ? 'ทั้งหมด' : $this->get_role_title($role));
      $expire_title = "อายุเอกสาร : ".($is_expired == 'all' ? 'ทั้งหมด' : ($is_expired == 1 ? 'เฉพาะที่หมดอายุ' : 'เฉพาะที่ไม่หมดอายุ'));
      $wh_title     = 'คลัง :  '. ($allWarehouse == 1 ? 'ทั้งหมด' : $this->get_title($warehouse));
      $ch_title = "ช่องทางขาย : ". ($allChannels == 1 ? 'ทั้งหมด' : $this->get_title($channels));
      $pay_title = "การชำระเงิน : ". ($allPayments == 1 ? "ทั้งหมด" : $this->get_title($payments));
      $state_title = "สถานะ : ".($allState == 1 ? "ทั้งหมด" : $this->get_state_title($state));

      //--- set report title header
      $sheet->setCellValue('A1', $report_title);
      $sheet->setCellValue('A2', $role_title);
      $sheet->setCellValue('A3', $expire_title);
      $sheet->setCellValue('A4', $state_title);
      $sheet->setCellValue('A5', $ch_title);
      $sheet->setCellValue('A6', $pay_title);
      $sheet->setCellValue('A7', $wh_title);
      $sheet->setCellValue('A8', 'วันที่เอกสาร : ('.thai_date($fromDate,'/') .') - ('.thai_date($toDate,'/').')');

      //--- set Table header
      $row = 9;

      $sheet->setCellValue("A{$row}", 'ลำดับ');
      $sheet->setCellValue("B{$row}", 'วันที่');
      $sheet->setCellValue("C{$row}", 'เลขที่');
      $sheet->setCellValue("D{$row}", 'อ้างอิง');
      $sheet->setCellValue("E{$row}", 'มูลค่า');
      $sheet->setCellValue("F{$row}", 'รหัสลูกค้า');
      $sheet->setCellValue("G{$row}", 'ชื่อลูกค้า');
      $sheet->setCellValue("H{$row}", 'สถานะ');
      $sheet->setCellValue("I{$row}", 'หมดอายุ');
      $sheet->setCellValue("J{$row}", 'วันที่แก้ไข');
      $sheet->setCellValue("K{$row}", 'ช่องทางขาย');
      $sheet->setCellValue("L{$row}", 'การชำระเงิน');
      $sheet->setCellValue("M{$row}", 'รหัสคลัง');
      $sheet->setCellValue("N{$row}", 'ชื่อคลัง');
      $sheet->setCellValue("O{$row}", 'Username');
      $sheet->setCellValue("P{$row}", 'พนักงาน');
      $sheet->setCellValue("Q{$row}", 'หมายเหตุ');
      $sheet->setCellValue("R{$row}", 'cancel reason');
      $sheet->setCellValue("S{$row}", 'Pre-order');
      $row++;

      //---- กำหนดความกว้างของคอลัมภ์
      $sheet->getColumnDimension('B')->setWidth(15);
      $sheet->getColumnDimension('C')->setWidth(20);
      $sheet->getColumnDimension('D')->setWidth(20);
      $sheet->getColumnDimension('E')->setWidth(20);
      $sheet->getColumnDimension('F')->setWidth(20);
      $sheet->getColumnDimension('G')->setWidth(90);
      $sheet->getColumnDimension('H')->setWidth(15);
      $sheet->getColumnDimension('I')->setWidth(10);
      $sheet->getColumnDimension('J')->setWidth(15);
      $sheet->getColumnDimension('K')->setWidth(20);
      $sheet->getColumnDimension('L')->setWidth(20);
      $sheet->getColumnDimension('M')->setWidth(15);
      $sheet->getColumnDimension('N')->setWidth(40);
      $sheet->getColumnDimension('O')->setWidth(15);
      $sheet->getColumnDimension('P')->setWidth(15);
      $sheet->getColumnDimension('Q')->setWidth(40);
      $sheet->getColumnDimension('R')->setWidth(40);
      $sheet->getColumnDimension('S')->setWidth(10);

      if( ! empty($details))
      {
        $state = $this->state_array();
        $wh = $this->warehouse_array();
        $channels = $this->channels_array();
        $payment = $this->payment_array();
        $cust = array();

        $no = 1;
        foreach($details as $rs)
        {          
          $date_add = Date::PHPToExcel($rs->date_add);
          $date_upd = Date::PHPToExcel($rs->date_upd);
          $amount = $this->order_details_model->get_doc_total($rs->code);

          if(empty($cust[$rs->customer_code]))
          {
            $customer_name = $this->customers_model->get_name($rs->customer_code);
            $cust[$rs->customer_code] = $customer_name;
          }

          $sheet->setCellValue("A{$row}", $no);
          $sheet->setCellValue("B{$row}", $date_add);
          $sheet->setCellValue("C{$row}", $rs->code);
          $sheet->setCellValue("D{$row}", $rs->reference);
          $sheet->setCellValue("E{$row}", $amount);
          $sheet->setCellValue("F{$row}", $rs->customer_code);
          $sheet->setCellValue("G{$row}", $cust[$rs->customer_code]);
          $sheet->setCellValue("H{$row}", $state[$rs->state]);
          $sheet->setCellValue("I{$row}", $rs->is_expired == 1 ? 'Y' : 'N');
          $sheet->setCellValue("J{$row}", $date_upd);
          $sheet->setCellValue("K{$row}", empty($channels[$rs->channels_code]) ? NULL : $channels[$rs->channels_code]);
          $sheet->setCellValue("L{$row}", empty($payment[$rs->payment_code]) ? NULL : $payment[$rs->payment_code]);
          $sheet->setCellValue("M{$row}", $rs->warehouse_code);
          $sheet->setCellValue("N{$row}", empty($wh[$rs->warehouse_code]) ? NULL : $wh[$rs->warehouse_code]);
          $sheet->setCellValue("O{$row}", $rs->user);
          $sheet->setCellValue("P{$row}", $this->user_model->get_name($rs->user));
          $sheet->setCellValue("Q{$row}", $rs->remark);

          if($rs->state == 9)
          {
            $sheet->setCellValue("R{$row}", $this->cancel_reason($rs->code));
          }

          $sheet->setCellValue("S{$row}", $rs->is_pre_order == 1 ? 'Y' : 'N');


          $no++;
          $row++;
        }

        $sheet->getStyle("B10:B{$row}")->getNumberFormat()->setFormatCode('dd/mm/yyyy');
        $sheet->getStyle("J10:J{$row}")->getNumberFormat()->setFormatCode('dd/mm/yyyy');
        $sheet->getStyle("E10:E{$row}")->getNumberFormat()->setFormatCode('#,##0.00');
      }
    }

    setToken($token);
    $file_name = "Report Order Details.xlsx";
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); /// form excel 2007 XLSX
    header('Content-Disposition: attachment;filename="'.$file_name.'"');
    $writer = new Xlsx($excel);
    $writer->save('php://output');

  }


  public function get_title($ds = array())
  {
    $list = "";
    if(!empty($ds))
    {
      $i = 1;
      foreach($ds as $rs)
      {
        $list .= $i === 1 ? $rs : ", {$rs}";
        $i++;
      }
    }

    return $list;
  }


  private function get_role_title($ds = array())
  {
    $list = "";

    $roles = array(
      'S' => 'WO',
      'C' => 'WC',
      'N' => 'WT',
      'P' => 'WS',
      'U' => 'WU',
      'T' => 'WQ',
      'Q' => 'WV',
      'L' => 'WL'
    );

    if( ! empty($ds))
    {
      $i = 1;
      foreach($ds as $rs)
      {
        if( ! empty($roles[$rs]))
        {
          $list .= $i=== 1 ? $roles[$rs] : ", {$roles[$rs]}";
          $i++;
        }
      }
    }

    return $list;
  }


  private function get_state_title($ds = array())
  {
    $list = "";

    $states = array(
      "1" => "รอดำเนินการ",
      "2" => "รอชำระเงิน",
      "3" => "รอจัดสินค้า",
      "4" => "กำลังจัดสินค้า",
      "5" => "รอตรวจ",
      "6" => "กำลังตรวจ",
      "7" => "รอเปิดบิล",
      "8" => "เปิดบิลแล้ว",
      "9" => "ยกเลิก"
    );

    if( ! empty($ds))
    {
      $i = 1;
      foreach($ds as $rs)
      {
        if( ! empty($states[$rs]))
        {
          $list .= $i === 1 ? $states[$rs] : ", {$states[$rs]}";
          $i++;
        }
      }
    }

    return $list;
  }
}
