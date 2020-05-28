package com.youyan.adapterrabbit.service;


import com.youyan.adapterrabbit.model.UdpData;
import org.springframework.amqp.core.AmqpAdmin;
import org.springframework.amqp.rabbit.core.RabbitTemplate;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import static com.youyan.adapterrabbit.constant.RabbitConstant.SENSOR_EXCHANGE;
import static com.youyan.adapterrabbit.constant.RabbitConstant.SENSOR_ROUTINGKEY;


@Service
public class RabbitService {

    @Autowired
    AmqpAdmin amqpAdmin;
    @Autowired
    RabbitTemplate rabbitTemplate;


    public void setUdpDataList(UdpData udpData){
        rabbitTemplate.convertAndSend(SENSOR_EXCHANGE,SENSOR_ROUTINGKEY,udpData);
    }


}
