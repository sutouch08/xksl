<?php
function employee_in($txt)
{
  $sc = array('0');
  $ci =& get_instance();
  $ci->load->model('masters/employee_model');
  $rs = $ci->employee_model->search($txt);

  if(!empty($rs))
  {
    foreach($rs as $cs)
    {
      $sc[] = $cs->empID;
    }
  }

  return $sc;
}
