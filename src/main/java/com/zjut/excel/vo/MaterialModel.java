package com.zjut.excel.vo;

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.Data;

/**
 * @author wangyong
 * @date 2021/5/8  15:44
 */
@Data
public class MaterialModel {

    @Excel(name = "供应商")
    private String supplier;

    @Excel(name = "序列号")
    private String serialNumber;

    @Excel(name = "物料编码")
    private String materialNo;

    @Excel(name = "图号")
    private String photoNo;

    @Excel(name = "名称")
    private String materialName;

    @Excel(name = "数量")
    private String number;

    @Excel(name = "项目名称")
    private String projectName;
}
