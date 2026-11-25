<?php
defined('BASEPATH') OR exit('No direct script access allowed');
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class Users extends PS_Controller
{
	public $menu_code = 'SCUSER'; //--- Add/Edit Users
	public $menu_group_code = 'SC'; //--- System security
	public $title = 'Users';

  public function __construct()
  {
    parent::__construct();
    $this->home = base_url().'users/users';
  }

  public function index()
  {
		$filter = array(
			'uname' => get_filter('uname', 'user', ''),
			'dname' => get_filter('dname', 'dname', ''),
			'profile' => get_filter('profile', 'profile', ''),
			'status' => get_filter('status', 'status', 'all')
		);

		$perpage = get_rows();
		$segment = 4; //-- url segment
		$rows = $this->user_model->count_rows($filter);
		$init	= pagination_config($this->home.'/index/', $rows, $perpage, $segment);
		$rs = $this->user_model->get_list($filter, $perpage, $this->uri->segment($segment));
		$filter['data'] = $rs;

		$this->pagination->initialize($init);
    $this->load->view('users/users_view', $filter);
  }





  public function add_user()
  {
		$this->load->helper('profile');
		$this->load->helper('saleman');
    $this->load->view('users/user_add_view');
  }


	public function edit_user($id)
	{
		$this->load->helper('profile');
		$this->load->helper('saleman');
		$ds['data'] = $this->user_model->get_user($id);
		$this->load->view('users/user_edit_view', $ds);
	}


	public function reset_password($id)
	{
			$this->title = 'Reset Password';
			$data['data'] = $this->user_model->get_user($id);
			$this->load->view('users/user_reset_pwd_view', $data);
	}



	public function change_password()
	{
		if($this->input->post('user_id'))
		{
			$id = $this->input->post('user_id');
			$pwd = password_hash($this->input->post('pwd'), PASSWORD_DEFAULT);
			$rs = $this->user_model->change_password($id, $pwd);

			if($rs === TRUE)
			{
				$arr = array(
					'last_pass_change' => date('Y-m-d')
				);
				//--- update last pass change
				$this->user_model->update_user($user->id, $arr);
				$this->session->set_flashdata('success', 'Password changed');
			}
			else
			{
				$this->session->set_flashdata('error', 'Change password not successfull, please try again');
			}
		}

		redirect($this->home);
	}



	public function delete_user($id)
	{
		$sc = TRUE;
		$user = $this->user_model->get_user($id);
		if(!empty($user))
		{
			if(!$this->user_model->has_transection($user->uname))
			{
				if(!$this->user_model->delete_user($id))
				{
					$sc = FALSE;
					$this->error = "Delete user failed";
				}
			}
			else
			{
				$sc = FALSE;
				$this->error = "ไม่สามารถลบ user ได้ เนื่องจากมี transection ในระบบแล้ว";
			}
		}
		else
		{
			$sc = FALSE;
			$this->error = "ไม่พบ User ที่ต้องการลบ";
		}

		echo $sc === TRUE ? 'success' : $this->error;
	}



	public function update_user()
	{
		$sc = TRUE;
		if($this->input->post('user_id'))
		{
			$id = $this->input->post('user_id');
			$uname = $this->input->post('uname');
			$dname = $this->input->post('dname');
			$id_profile = $this->input->post('profile') === '' ? NULL : $this->input->post('profile');
			$sale_id = $this->input->post('sale_id') === '' ? NULL : $this->input->post('sale_id');
			$status = $this->input->post('status');
			$is_viewer = $this->input->post('is_viewer');

			$data = array(
				'uname' => $uname,
				'name' => $dname,
				'id_profile' => $id_profile,
				'sale_id' => $sale_id,
				'active' => $status,
				'is_viewer' => $is_viewer
			);

			$rs = $this->user_model->update_user($id, $data);
			if($rs === FALSE)
			{
				$this->session->set_flashdata('error', 'Update user not successfully');
			}
			else
			{
				$this->session->set_flashdata('success', 'User updated');
			}
		}
		else
		{
			$this->session->set_flashdata('error','Update fail : data not found');
		}

		redirect($this->home.'/edit_user/'.$id);

	}




	public function new_user()
	{
		if($this->input->post('uname'))
		{
			$uname = $this->input->post('uname');
			$dname = $this->input->post('dname');
			$pwd = password_hash($this->input->post('pwd'), PASSWORD_DEFAULT);
			$uid = md5(uniqid());
			$id_profile = $this->input->post('profile') === '' ? NULL : $this->input->post('profile');
			$sale_id = $this->input->post('sale_id') === '' ? NULL : $this->input->post('sale_id');
			$status = $this->input->post('status');
			$is_viewer = $this->input->post('is_viewer');

			$data = array(
				'uname' => $uname,
				'pwd' => $pwd,
				'name' => $dname,
				'uid' => $uid,
				'id_profile' => $id_profile,
				'sale_id' => $sale_id,
				'active' => $status,
				'is_viewer' => $is_viewer,
				'last_pass_change' => date('Y-m-d')
			);

			$rs = $this->user_model->new_user($data);

			if($rs === FALSE)
			{
				set_error('Create User fail');
			}
			else
			{
				set_message('User created');
			}

		}
		else
		{

			set_error('Create User fail : Empty data');
		}

		redirect($this->home.'/add_user');
	}




	public function valid_dname($dname, $id = '')
	{
		$rs = $this->user_model->is_exists_display_name($dname, $id);

		if($rs === TRUE)
		{
			echo 'exists';
		}
		else
		{
			echo 'not exists';
		}
	}



	public function valid_uname($uname, $id = '')
	{
		$rs = $this->user_model->is_exists_uname($uname, $id);
		if($rs === TRUE)
		{
			echo 'exists';
		}
		else
		{
			echo 'not exists';
		}
	}




	//--- Activeate suspend user by id;
	public function active_user($id)
	{
		$rs = $this->user_model->active_user($id);
		echo $rs === TRUE ? 'success' : json_encode($rs);
	}






	//--- Suspend activated user by id
	public function disactive_user($id)
	{
		$rs = $this->user_model->disactive_user($id);

		echo $rs === TRUE ? 'success' : $rs;
	}


	public function get_user_permissions($id)
	{
		$this->load->model('users/permission_model');
		$sc = TRUE;
		$ds = array();

		$user = $this->user_model->get_user($id);

		if( ! empty($user))
		{
			$ds['header'] = "Permission : \"{$user->uname}\"";
			$ds['group'] = array();

			$groups = $this->menu->get_active_menu_groups();

			if( ! empty($groups))
			{
				foreach($groups as $gp)
				{
					if($gp->pm)
					{
						$menuGroup = array(
							'group_code' => $gp->code,
							'group_name' => $gp->name,
							'menu' => ''
						);

						$menus = $this->menu->get_menus_by_group($gp->code);

						if( ! empty($menus))
						{
							$item = array();

							foreach($menus as $menu)
							{
								if($menu->valid)
								{
									$pm = $this->permission_model->get_permission($menu->code, $user->id_profile);

									$arr = array(
										'menu_code' => $menu->code,
										'menu_name' => $menu->name,
										'cv' => $pm->can_view ? 1 : 0,
										'ca' => $pm->can_add ? 1 : 0,
										'ce' => $pm->can_edit ? 1 : 0,
										'cd' => $pm->can_delete ? 1 : 0,
										'cp' => $pm->can_approve ? 1 : 0
									);

									array_push($item, $arr);
								}

							}

							$menuGroup['menu'] = $item;
						}

						array_push($ds['group'], $menuGroup);
					}
				}
			}
		}
		else
		{
			$sc = FALSE;
			$this->error = "Invalid user id";
		}

		echo $sc === TRUE ? json_encode($ds) : $this->error;
	}




	public function export_permission()
	{
		$this->load->model('users/permission_model');
		$this->load->model('users/profile_model');
		$token = $this->input->post('token');
		$id = $this->input->post('user_id');

		$user = $this->user_model->get_user($id);
		$uname = empty($user) ? 'no data' : $user->uname;

    $excel = new Spreadsheet();
		$sheet = $excel->getActiveSheet();
    $excel->setActiveSheetIndex(0);
    $sheet->setTitle($uname);

		if( ! empty($user))
		{
			$sheet->getColumnDimension('A')->setWidth('30');
			$sheet->getColumnDimension('B')->setWidth('15');
			$sheet->getColumnDimension('C')->setWidth('15');
			$sheet->getColumnDimension('D')->setWidth('15');
			$sheet->getColumnDimension('E')->setWidth('15');
			$sheet->getColumnDimension('F')->setWidth('15');

			$sheet->setCellValue('A1', 'User : ')->getStyle('A1')->getAlignment()->setHorizontal('right');
			$sheet->setCellValue('B1', $user->uname);
			$sheet->mergeCells('B1:C1');

			$sheet->setCellValue('D1', 'Display name : ')->getStyle('D1')->getAlignment()->setHorizontal('right');
			$sheet->setCellValue('E1', $user->name);
			$sheet->mergeCells('E1:F1');

			$sheet->setCellValue('A2', 'Profile : ')->getStyle('A2')->getAlignment()->setHorizontal('right');
			$sheet->setCellValue('B2', $this->profile_model->get_name($user->id_profile));
			$sheet->mergeCells('B2:C2');
			$sheet->setCellValue('D2', "Status : ")->getStyle('D2')->getAlignment()->setHorizontal('right');
			$sheet->setCellValue('E2', ($user->active == 1 ? 'Active' : 'Inactive'));
			$sheet->mergeCells('E2:F2');

			$row = 4;


			$groups = $this->menu->get_active_menu_groups();

			if( ! empty($groups))
			{
				foreach($groups as $gp)
				{
					if($gp->pm)
					{
						$sheet->setCellValue("A{$row}", $gp->name);
						$sheet->setCellValue("B{$row}", 'ดู');
						$sheet->setCellValue("C{$row}", 'เพิ่ม');
						$sheet->setCellValue("D{$row}", 'แก้ไข');
						$sheet->setCellValue("E{$row}", 'ลบ');
						$sheet->setCellValue("F{$row}", 'อนุมัติ');

						$sheet->getStyle("A{$row}:F{$row}")
							->getFill()
							->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
							->getStartColor()
							->setARGB('F28A8C');
						$row++;

						$menus = $this->menu->get_menus_by_group($gp->code);

						if( ! empty($menus))
						{
							foreach($menus as $menu)
							{
								if($menu->valid)
								{
									$pm = $this->permission_model->get_permission($menu->code, $user->id_profile);

									$sheet->setCellValue("A{$row}", $menu->name);
									$sheet->setCellValue("B{$row}", ($pm->can_view ? 'Y' : '-'));
									$sheet->setCellValue("C{$row}", ($pm->can_add ? 'Y' : '-'));
									$sheet->setCellValue("D{$row}", ($pm->can_edit ? 'Y' : '-'));
									$sheet->setCellValue("E{$row}", ($pm->can_delete ? 'Y' : '-'));
									$sheet->setCellValue("F{$row}", ($pm->can_approve ? 'Y' : '-'));

									$row++;
								}
							}
						}
					} //-- endif
				} //--- end foreach

				if($row > 3)
				{
					$sheet->getStyle("B3:F{$row}")->getAlignment()->setHorizontal('center');
				}
			} //--- endif group
		}

		setToken($token);
    $file_name = "{$uname} Permission.xlsx";
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); /// form excel 2007 XLSX
    header('Content-Disposition: attachment;filename="'.$file_name.'"');
    $writer = new Xlsx($excel);
    $writer->save('php://output');
	}



	public function export_all_permission()
	{
		ini_set('memory_limit', '2048M');
		ini_set('max_execution_time', 600);

		$this->load->model('users/permission_model');
		$this->load->model('users/profile_model');
		$token = $this->input->post('alltoken');
		$all = $this->input->post('all') == 1 ? TRUE : FALSE;
		$users = $this->user_model->get_all($all);

		$ds = array();

		$groups = $this->menu->get_active_menu_groups();

		if (! empty($groups))
		{
			foreach ($groups as $group)
			{
				if ($group->pm)
				{
					$arr = array(
						'name' => $group->name,
						'menus' => NULL
					);

					$menus = $this->menu->get_menus_by_group($group->code);

					if (! empty($menus))
					{
						$items = array();

						foreach ($menus as $menu)
						{
							if ($menu->valid)
							{
								$items[] = array(
									'code' => $menu->code,
									'name' => $menu->name
								);
							}
						}

						$arr['menus'] = $items;
					}

					$ds[] = $arr;
				}
			}
		}

		$excel = new Spreadsheet();

		if (! empty($users))
		{
			$index = 0;

			foreach ($users as $user)
			{
				$sheet = $index === 0 ? $excel->getActiveSheet() : $excel->createSheet();
				$sheet->setTitle($user->uname);
				$tabColor = $user->active == 1 ? '54c784' : 'c96b65';
				$sheet->getTabColor()->setARGB($tabColor);

				$sheet->getColumnDimension('A')->setWidth('30');
				$sheet->getColumnDimension('B')->setWidth('15');
				$sheet->getColumnDimension('C')->setWidth('15');
				$sheet->getColumnDimension('D')->setWidth('15');
				$sheet->getColumnDimension('E')->setWidth('15');
				$sheet->getColumnDimension('F')->setWidth('15');

				$sheet->setCellValue('A1', 'User : ')->getStyle('A1')->getAlignment()->setHorizontal('right');
				$sheet->setCellValue('B1', $user->uname);
				$sheet->mergeCells('B1:C1');

				$sheet->setCellValue('D1', 'Display name : ')->getStyle('D1')->getAlignment()->setHorizontal('right');
				$sheet->setCellValue('E1', $user->name);
				$sheet->mergeCells('E1:F1');

				$sheet->setCellValue('A2', 'Profile : ')->getStyle('A2')->getAlignment()->setHorizontal('right');
				$sheet->setCellValue('B2', $this->profile_model->get_name($user->id_profile));
				$sheet->mergeCells('B2:C2');
				$sheet->setCellValue('D2', "Status : ")->getStyle('D2')->getAlignment()->setHorizontal('right');
				$sheet->setCellValue('E2', ($user->active == 1 ? 'Active' : 'Inactive'));
				$sheet->mergeCells('E2:F2');

				$row = 4;

				if (! empty($ds))
				{
					foreach ($ds as $rs)
					{
						$sheet->setCellValue("A{$row}", $rs['name']);
						$sheet->setCellValue("B{$row}", 'ดู');
						$sheet->setCellValue("C{$row}", 'เพิ่ม');
						$sheet->setCellValue("D{$row}", 'แก้ไข');
						$sheet->setCellValue("E{$row}", 'ลบ');
						$sheet->setCellValue("F{$row}", 'อนุมัติ');

						$sheet->getStyle("A{$row}:F{$row}")
							->getFill()
							->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
							->getStartColor()
							->setARGB('F28A8C');

						$row++;

						$menus = $rs['menus'];

						if (! empty($menus))
						{
							foreach ($menus as $menu)
							{
								$pm = $this->permission_model->get_permission($menu['code'], $user->id_profile);

								$sheet->setCellValue("A{$row}", $menu['name']);
								$sheet->setCellValue("B{$row}", ($pm->can_view ? 'Y' : '-'));
								$sheet->setCellValue("C{$row}", ($pm->can_add ? 'Y' : '-'));
								$sheet->setCellValue("D{$row}", ($pm->can_edit ? 'Y' : '-'));
								$sheet->setCellValue("E{$row}", ($pm->can_delete ? 'Y' : '-'));
								$sheet->setCellValue("F{$row}", ($pm->can_approve ? 'Y' : '-'));

								$row++;
							}
						}
					} //--- end foreach

					if ($row > 4)
					{
						$sheet->getStyle("B3:F{$row}")->getAlignment()->setHorizontal('center');
					}
				} //--- endif group

				$index++;
			}
		}

		setToken($token);
		$file_name = "Users Permission.xlsx";
		header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'); /// form excel 2007 XLSX
		header('Content-Disposition: attachment;filename="' . $file_name . '"');
		$writer = new Xlsx($excel);
		$writer->save('php://output');
	}



	public function clear_filter()
	{
		$filter = array('user', 'dname', 'profile');
		clear_filter($filter);
		echo 'done';
	}

}//--- end class


 ?>
