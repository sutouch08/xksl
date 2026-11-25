<?php

function get_available_stock($item)
{
  $ci =& get_instance();
  $ci->load->model('stock/stock_model');
  $ci->load->model('orders/orders_model');
  $sell_stock = $ci->stock_model->get_sell_stock($item);
  $reserv_stock = $ci->orders_model->get_reserv_stock($item);
  $availableStock = $sell_stock - $reserv_stock;
  return $availableStock < 0 ? 0 : $availableStock;
}