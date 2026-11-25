<?php

function select_saleman($id = "")
{
  $ci =& get_instance();
  $ci->load->model('masters/slp_model');
  $result = $ci->slp_model->get_data();
  $ds = '';

  if(!empty($result))
  {
    foreach($result as $rs)
    {
      $ds .= '<option value="'.$rs->id.'" '.is_selected($rs->id, $id).'>'.$rs->name.'</option>';
    }
  }

  return $ds;
}