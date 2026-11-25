<?php

function discount_policy_in($txt)
{
  $sc = "0";
  $ci =& get_instance();
  $ci->load->model('discount/discount_policy_model');
  $rs = $ci->discount_policy_model->search($txt);

  if(!empty($rs))
  {
    foreach($rs as $cs)
    {
      $sc .= ", ".$cs->id;
    }
  }

  return $sc;
}
