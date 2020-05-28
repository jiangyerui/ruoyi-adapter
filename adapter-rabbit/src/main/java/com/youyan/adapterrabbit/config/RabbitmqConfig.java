package com.youyan.adapterrabbit.config;

import org.springframework.amqp.core.Binding;
import org.springframework.amqp.core.BindingBuilder;
import org.springframework.amqp.core.DirectExchange;
import org.springframework.amqp.core.Queue;
import org.springframework.amqp.support.converter.Jackson2JsonMessageConverter;
import org.springframework.amqp.support.converter.MessageConverter;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

import static com.youyan.adapterrabbit.constant.RabbitConstant.*;


@Configuration
public class RabbitmqConfig {
    /**
     * 配置JSON转换器，自动JSON保存，自动JSON取出
     */
    @Bean
    public MessageConverter messageConverter(){
        return new Jackson2JsonMessageConverter();
    }

    //队列 起名：directQueue
    @Bean
    public Queue directQueue() {
        return new Queue(SENSOR_QUEUE,true);  //true 是否持久
    }

    //Direct交换机 起名：directExchange
    @Bean
    DirectExchange directExchange() {
        return new DirectExchange(SENSOR_EXCHANGE);
    }

    //绑定  将队列和交换机绑定, 并设置用于匹配键：TestDirectRouting
    @Bean
    Binding bindingDirect() {
        return BindingBuilder.bind(directQueue()).to(directExchange()).with(SENSOR_ROUTINGKEY);
    }

}
