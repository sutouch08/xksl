<?php
function dispatch_status($status)
{
  $ds = array(
    'P' => 'Pending',
    'R' => 'Released',
    'S' => 'Shipped',
    'C' => 'Closed',
    'D' => 'Canceled'
  );

  return empty($ds[$status ?? '']) ? 'Unknow' : $ds[$status];
}
