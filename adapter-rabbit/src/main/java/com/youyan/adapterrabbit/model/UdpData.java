package com.youyan.adapterrabbit.model;

import java.io.Serializable;
import java.util.Arrays;

public class UdpData implements Serializable {
    private byte[] req;
    private String hostName;

    public UdpData() {
    }

    public UdpData(byte[] req) {
        this.req = req;
    }

    public UdpData(byte[] req, String hostName) {
        this.req = req;
        this.hostName = hostName;
    }

    public byte[] getReq() {
        return req;
    }

    public void setReq(byte[] req) {
        this.req = req;
    }

    public String getHostName() {
        return hostName;
    }

    public void setHostName(String hostName) {
        this.hostName = hostName;
    }

    @Override
    public String toString() {
        return "UdpData{" +
                "req=" + Arrays.toString(req) +
                ", hostName='" + hostName + '\'' +
                '}';
    }
}
