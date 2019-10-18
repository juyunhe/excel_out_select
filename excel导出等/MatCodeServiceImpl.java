package com.richfit.sod.appserver.service.resource.impl;

import cn.afterturn.easypoi.excel.entity.TemplateExportParams;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.github.pagehelper.PageHelper;
import com.github.pagehelper.PageInfo;
import com.richfit.sod.appserver.common.BaseResult;
import com.richfit.sod.appserver.common.MessageConfig;
import com.richfit.sod.appserver.dao.resource.MatCodeDao;
import com.richfit.sod.appserver.entity.MatCode;
import com.richfit.sod.appserver.entity.SysUser;
import com.richfit.sod.appserver.exceptions.BaseException;
import com.richfit.sod.appserver.service.resource.MatCodeService;
import com.richfit.sod.appserver.utils.*;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.*;
import java.util.stream.Collectors;

/**
 * 应急资源编码接口实现类
 */
@Service
public class MatCodeServiceImpl implements MatCodeService {

    private Logger logger = LoggerFactory.getLogger(this.getClass());

    @Autowired
    private MatCodeDao matCodeDao;

    /**
     * 根据 id 获取应急资源编码详细信息
     *
     * @param resourceCodeId
     * @return
     */
    @Override
    public MatCode findMatCodeById(String resourceCodeId) {
        //应急资源编码id不能为空
        ObjectUtil.isEmptyWithMsgCode(resourceCodeId, "SELECT_MAT_CODE_FIAL");
        MatCode matCode = null;
        try {
            matCode = matCodeDao.selectById(resourceCodeId);

        } catch (Exception e) {
            logger.error("根据 id 获取应急资源编码详细信息失败:" + e.getMessage(), e);
            ObjectUtil.throwErrorWithMsgCode("MAT_CODE_FIND_ID_ERROR");
        }
        return matCode;
    }

    /**
     * 保存或修改应急资源编码信息
     *
     * @param matCode
     */
    @Transactional(rollbackFor = Exception.class, value = "dataSourceManager")
    @Override
    public void addOrUpdateMatCode(MatCode matCode) {
        //应急资源编码必填
        ObjectUtil.isEmptyWithMsgCode(matCode.getResourceCode(), "THE_RESOURCE_CODE_IS_NOT_NULL");
        //应急资源计量单位必填-去掉
        // ObjectUtil.isEmptyWithMsgCode(matCode.getUnit(),"THE_UNIT_IS_NOT_NULL");
        //资源名称必填
        ObjectUtil.isEmptyWithMsgCode(matCode.getResourceName(), "RESOURCE_NAME_IS_NOT_NULL");
        //应急资源类别必填
        ObjectUtil.isEmptyWithMsgCode(matCode.getResourceType(), "RESOURCE_TYPE_IS_NOT_NULL");
        //规格型号必填
        ObjectUtil.isEmptyWithMsgCode(matCode.getSpecificationModel(), "SPECIFICATION_MODEL_IS_NOT_NULL");
        if (StringUtils.isEmpty(matCode.getResourceCodeId())) {
            try {
                //保存操作
                matCode.setResourceCodeId(UuidUtil.getUuid());
                ControllerUtil.addMustFields(matCode);
                matCodeDao.insertMatCode(matCode);
            } catch (Exception e) {
                logger.error("保存应急资源编码信息失败:" + e.getMessage(), e);
                ObjectUtil.throwErrorWithMsgCode("MAT_CODE_ADD_ERROR");
            }
        } else {
            try {
                ControllerUtil.updateMustFields(matCode);
                //修改操作
                matCodeDao.updateMatCode(matCode);
            } catch (Exception e) {
                logger.error("修改应急资源编码信息失败:" + e.getMessage(), e);
                ObjectUtil.throwErrorWithMsgCode("MAT_CODE_UPDATE_ERROR");
            }
        }
    }

    /**
     * 删除应急资源编码信息
     *
     * @param id
     */
    @Transactional(rollbackFor = Exception.class, value = "dataSourceManager")
    @Override
    public void deleteMatCode(String id) {
        ObjectUtil.isEmptyWithMsgCode(id, "FAIL_TO_DELETE");
        try {
            MatCode matCode = new MatCode();
            matCode.setResourceCodeId(id);
            ControllerUtil.deleteMustFields(matCode);
            matCodeDao.deleteMatCode(matCode);
        } catch (Exception e) {
            logger.error("删除应急资源编码信息失败:" + e.getMessage(), e);
            ObjectUtil.throwErrorWithMsgCode("FAIL_TO_DELETE");
        }
    }

    /**
     * 分页查看根据用户权限获取全部的应急资源编码信息
     *
     * @param pageIndex
     * @param pageSize
     * @return
     */
    @Override
    public Map<String, Object> findMatCode(Integer pageIndex, Integer pageSize) {
        Map<String, Object> map = new HashMap<>();
        try {
            PageHelper.startPage(pageIndex, pageSize);
            List<MatCode> matCodes = matCodeDao.findMatCode();
            PageInfo page = new PageInfo(matCodes, pageSize);
            long totalSize = page.getTotal();
            int totalPages = page.getPages();
            map.put("totalSize", totalSize);
            map.put("totalPages", totalPages);
            map.put("data", matCodes);
        } catch (Exception e) {
            logger.error("分页查看根据用户权限获取全部的应急资源编码信息失败:" + e.getMessage(), e);
            ObjectUtil.throwErrorWithMsgCode("MAT_CODE_FIND_ERROR");
        }
        return map;
    }

    /**
     * 分页查看根据关键字查询匹配的应急资源编码信息
     *
     * @param pageIndex
     * @param pageSize
     * @param matCode
     * @return
     */
    @Override
    public Map<String, Object> searchMatCode(Integer pageIndex, Integer pageSize, MatCode matCode) {
        Map<String, Object> map = new HashMap<>();
        List<MatCode> newMatCodes = null;
        try {
            PageHelper.startPage(pageIndex, pageSize);
            List<MatCode> matCodes = matCodeDao.searchMatCode(matCode);
            //获取编码对应的名称
            String[] fields = {"resourceType"};
            String[] codeTypes = {"SODEMED07"};
            String[] fieldNames = {"resourceTypeName"};
            newMatCodes = CodeUtil.transitionArray(matCodes, MatCode.class, fields, codeTypes, fieldNames, MessageConfig.getMessageLanguage());
            PageInfo page = new PageInfo(matCodes, pageSize);
            long totalSize = page.getTotal();
            int totalPages = page.getPages();
            map.put("totalSize", totalSize);
            map.put("totalPages", totalPages);
            map.put("data", newMatCodes);
        } catch (Exception e) {
            logger.error("分页查看根据关键字查询匹配的应急资源编码信息失败:" + e.getMessage(), e);
            ObjectUtil.throwErrorWithMsgCode("MAT_CODE_SEARCH_ERROR");
        }
        return map;
    }

    /**
     * 根据获取到的应急资源编码信息进行去重
     *
     * @param maps
     * @return
     */
    @Override
    public Map<String, Object> findMatCodeByDelRepetition(Map<String, Object> maps) {
        Map<String, Object> map = new HashMap<>();
        try {
            List<String> resourceCodes = (List<String>) maps.get("resourceCodes");
            Integer pageIndex = Integer.valueOf(maps.get("pageIndex").toString());
            Integer pageSize = Integer.valueOf(maps.get("pageSize").toString());
            String resourceName = (String) maps.get("resourceName");
            String resourceCode = (String) maps.get("resourceCode");
            MatCode matCode = new MatCode();
            if (StringUtils.isNotEmpty(resourceName)) {
                matCode.setResourceName(resourceName);
            }
            if (StringUtils.isNotEmpty(resourceCode)) {
                matCode.setResourceCode(resourceCode);
            }
            PageHelper.startPage(pageIndex, pageSize);
            List<MatCode> matCodes = matCodeDao.searchMatCode(matCode);
            //获取编码对应的名称
            String[] fields = {"resourceType"};
            String[] codeTypes = {"SODEMED07"};
            String[] fieldNames = {"resourceTypeName"};
            //List<MatCode> newMatCodes;
            matCodes = CodeUtil.transitionArray(matCodes, MatCode.class, fields, codeTypes, fieldNames, MessageConfig.getMessageLanguage());
            //如果没有传回的编码则将所有的数据全部返回，否则进行去重返回
            if (Objects.nonNull(resourceCodes) && resourceCodes.size() != 0) {
                matCodes = matCodes.stream().filter(item -> {
                    return !resourceCodes.contains(item.getResourceCode());
                }).collect(Collectors.toList());
            } else {
                matCodes = matCodes;
            }
            PageInfo page = new PageInfo(matCodes, pageSize);
            long totalSize = page.getTotal();
            int totalPages = page.getPages();
            map.put("totalSize", totalSize);
            map.put("totalPages", totalPages);
            map.put("data", matCodes);
        } catch (Exception e) {
            logger.error("根据获取到的应急资源编码信息进行去重失败:" + e.getMessage(), e);
            ObjectUtil.throwErrorWithMsgCode("MAT_CODE_DEL_REPETITION_ERROR");
        }
        return map;
    }


    /**
     * 限制应急资源编码不能重复
     *
     * @param resourceCode
     */
    @Override
    public String resourceCodeNotRepeat(String resourceCode, String resourceCodeId) {
        try {
            if (StringUtils.isNotEmpty(resourceCode)) {
                //限制应急资源编码不能重复
                MatCode newMatCode = matCodeDao.selectResourceCodeRepetition(resourceCode, resourceCodeId);
                if (Objects.nonNull(newMatCode)) {
                    if (StringUtils.isNotEmpty(newMatCode.getResourceCode())) {
                        //限制应急资源编码不能重复 Limit emergency resource coding cannot be repeated!
                        //ObjectUtil.throwErrorWithMsgCode("THE_RESOURCE_CODE_REPEAT");
                        return newMatCode.getResourceCode();
                    }
                }
            }
        } catch (BaseException e) {
            //限制应急资源编码不能重复 Limit emergency resource coding cannot be repeated!
            ObjectUtil.throwErrorWithMsgCode("THE_RESOURCE_CODE_REPEAT");
        }
        return null;
    }

    /**
     * 根据资源存储点id获取指定的应急编码
     *
     * @param maps
     * @return
     */
    @Override
    public Map<String, Object> findMatCodeByWarehouse(Map<String, Object> maps) {
        Map<String, Object> map = new HashMap<>();
        try {
            Integer pageIndex = Integer.valueOf(maps.get("pageIndex").toString());
            Integer pageSize = Integer.valueOf(maps.get("pageSize").toString());
            String resourceName = (String) maps.get("resourceName");
            String resourceCode = (String) maps.get("resourceCode");
            PageHelper.startPage(pageIndex, pageSize);
            MatCode matCode = new MatCode();
            if (StringUtils.isNotEmpty(resourceName))
                matCode.setResourceName(resourceName);
            if (StringUtils.isNotEmpty(resourceCode))
                matCode.setResourceCode(resourceCode);
            List<MatCode> matCodes = matCodeDao.searchMatCode(matCode);
            //获取编码对应的名称
            String[] fields = {"resourceType"};
            String[] codeTypes = {"SODEMED07"};
            String[] fieldNames = {"resourceTypeName"};
            matCodes = CodeUtil.transitionArray(matCodes, MatCode.class, fields, codeTypes, fieldNames, MessageConfig.getMessageLanguage());
            PageInfo page = new PageInfo(matCodes, pageSize);
            long totalSize = page.getTotal();
            int totalPages = page.getPages();
            map.put("totalSize", totalSize);
            map.put("totalPages", totalPages);
            map.put("data", matCodes);
        } catch (Exception e) {
            logger.error("根据获取到的应急资源编码信息进行去重失败:" + e.getMessage(), e);
            ObjectUtil.throwErrorWithMsgCode("MAT_CODE_DEL_REPETITION_ERROR");
        }
        return map;
    }

    /**
     * 应急资源编码信息模板下载
     * @param response
     */
    @Override
    public void matCodeTemplate(HttpServletResponse response) {
        try {
            //获取所有应急编码类型（CodeByType）
            JSONArray json = CodeUtil.getMultilevelCodeByType("SODEMED07", MessageConfig.getMessageLanguage());
            String[] resourceTypes = new String[json.size()];
            for (int i = 0; i < json.size(); i++) {
                JSONObject jsonObject = json.getJSONObject(i);
                resourceTypes[i] = (String) jsonObject.get("label");
            }
            //获取标题
            String title = MessageConfig.getMessage("RESOURCE_CODE_TITLE");
            //获取表头
            String[] codeHeader = MessageConfig.getMessage("MAT_CODE_HEADER").split(",");
            codeHeader[0] =  " "+codeHeader[0] ;
            codeHeader[1] =  "*"+codeHeader[1] ;
            codeHeader[2] =  "*"+codeHeader[2] ;
            codeHeader[3] =  "*"+codeHeader[3] ;
            codeHeader[4] =  "*"+codeHeader[4] ;
            codeHeader[5] =  " "+codeHeader[5] ;
            codeHeader[6] =  " "+codeHeader[6] ;
            List<String[]> lists = new ArrayList<>();
            lists.add(resourceTypes);
            int[] naturalColumnIndexs = {3};
            //导出
            //创建模板
            HSSFWorkbook hssfWorkbook = ExportExcelTemplateUtil.getHSSFWorkbook(title, codeHeader, null, lists, naturalColumnIndexs);

            //更改请求头
            ExcelUtiles.downLoadExcel("resourceCodeTemplate.xls", response, hssfWorkbook);
        } catch (Exception e) {
            logger.error("应急资源编码信息模板下载失败:" + e.getMessage(), e);
            ObjectUtil.throwErrorWithMsgCode("MAT_CODE_EXPORT_TEMPLATE_ERROR");
        }
    }

    /**
     * 应急资源编码信息导出
     *
     * @param code
     * @param request
     * @param response
     */
    @Override
    public void matCodeExport(MatCode code, HttpServletRequest request, HttpServletResponse response) {
        try {
            List<MatCode> matCodes = matCodeDao.searchMatCode(code);
            //获取编码对应的名称
            String[] fields = {"resourceType"};
            String[] codeTypes = {"SODEMED07"};
            String[] fieldNames = {"resourceTypeName"};
            matCodes = CodeUtil.transitionArray(matCodes, MatCode.class, fields, codeTypes, fieldNames, MessageConfig.getMessageLanguage());
            Map<String, Object> map1 = new HashMap<>();
            List<Map<String, String>> list = new ArrayList<>();
            int i = 1;
            for (MatCode item : matCodes) {
                Map<String, String> maps = new HashMap<>();
                maps.put("index", Integer.toString(i));
                maps.put("resourceCode", item.getResourceCode());
                maps.put("resourceType", item.getResourceTypeName());
                maps.put("resourceName", item.getResourceName());
                maps.put("specificationModel", item.getSpecificationModel());
                maps.put("manufacturer", item.getManufacturer());
                maps.put("unit", item.getUnit());
                list.add(maps);
                i++;
            }
            String title = MessageConfig.getMessage("RESOURCE_CODE_TITLE");
            map1.put("title", title);
            String[] codeHeader = MessageConfig.getMessage("MAT_CODE_HEADER").split(",");
            Map<String, String> mapHeader = new HashMap<>();
            mapHeader.put("index", codeHeader[0]);
            mapHeader.put("resourceCode", codeHeader[1]);
            mapHeader.put("resourceType", codeHeader[2]);
            mapHeader.put("resourceName", codeHeader[3]);
            mapHeader.put("specificationModel", codeHeader[4]);
            mapHeader.put("manufacturer", codeHeader[5]);
            mapHeader.put("unit", codeHeader[6]);
            map1.put("header", mapHeader);
            map1.put("data", list);
            map1.put("operatorName", MessageConfig.getMessage("OPERATOR"));
            //获取操作人
            SysUser user = AuthorityUtil.getCurrentUser();
            map1.put("operator", user.getUserName());
            TemplateExportParams params = new TemplateExportParams(
                    ExcelUtiles.convertTemplatePath("static/excelTemplate/resource/resourceCode.xlsx"));
            ExcelUtiles.exportTemplateExcel("resourceCode.xlsx", params, map1,
                    MessageConfig.getMessage("SHEET_CODE_NAME"), response);
        } catch (Exception e) {
            logger.error("应急资源编码信息导出失败:" + e.getMessage(), e);
            ObjectUtil.throwErrorWithMsgCode("MAT_CODE_EXPORT_ERROR");
        }
    }

    /**
     * 应急资源编码信息导入
     *
     * @param file
     */
    @Override
    public BaseResult importExcel(MultipartFile file) {
        //判断异常标志位
        boolean isError = false;
        boolean isAdd = true;
        StringBuffer buffer = new StringBuffer();
        try {
            List<MatCode> matCodes = new ArrayList<>();
            //获取导入的数据
            List<Map> maps = ExcelUtiles.importExcel(file, 0, 2, Map.class);
            //获取到表头必须按顺序来接收
            String[] codeHeaders = MessageConfig.getMessage("MAT_CODE_HEADER").split(",");
            /*
             * 导入验证步骤：
             * 1、校验每一行必填项是否已经填写，如果没有填写，记录在返回的错误提示消息中，提示内容：第几行哪一列值为空
             * 2、必填项校验通过后进行如下校验：
             *    1）、校验Resource code是否已经存在，如果存在，也需要记录再返回的错误提示消息中，提示内容：第几行Resource code已存在，已找到对应资源编码信息
             *    2）、编码转换，根据编码英文值去找对应code
             * */
            for (Map map : maps) {
                String resourceType = (String) map.get(codeHeaders[2]);
                if (StringUtils.isNotBlank(resourceType)) {
                    //获取到所有的编码类型
                    JSONArray json = CodeUtil.getMultilevelCodeByType(
                            "SODEMED07", MessageConfig.getMessageLanguage());
                    //如果编码名称相等则替换为对应的编码id
                    for (int i = 0; i < json.size(); i++) {
                        JSONObject jsonObject = json.getJSONObject(i);
                        if (resourceType.equals((String) jsonObject.get("label"))) {
                            resourceType = (String) jsonObject.get("dictCode");
                        }
                    }
                }
                //获取导入的数据到对象
                MatCode matCode = new MatCode();
                if (map.get(codeHeaders[1]) instanceof Integer) {
                    matCode.setResourceCode(Integer.toString((Integer) map.get(codeHeaders[1])));
                } else {
                    matCode.setResourceCode((String) map.get(codeHeaders[1]));
                }
                matCode.setResourceType(resourceType);
                if (map.get(codeHeaders[3]) instanceof Integer) {
                    matCode.setResourceName(Integer.toString((Integer) map.get(codeHeaders[3])));
                } else {
                    matCode.setResourceName((String) map.get(codeHeaders[3]));
                }
                if (map.get(codeHeaders[4]) instanceof Integer) {
                    matCode.setSpecificationModel(Integer.toString((Integer) map.get(codeHeaders[4])));
                } else {
                    matCode.setSpecificationModel((String) map.get(codeHeaders[4]));
                }
                if (map.get(codeHeaders[5]) instanceof Integer) {
                    matCode.setManufacturer(Integer.toString((Integer) map.get(codeHeaders[5])));
                } else {
                    matCode.setManufacturer((String) map.get(codeHeaders[5]));
                }
                if (map.get(codeHeaders[6]) instanceof Integer) {
                    matCode.setUnit(Integer.toString((Integer) map.get(codeHeaders[6])));
                } else {
                    matCode.setUnit((String) map.get(codeHeaders[6]));
                }
                //验证编码不重复
                //限制应急资源编码不能重复
                MatCode newMatCode = matCodeDao.selectResourceCodeRepetition(matCode.getResourceCode(), matCode.getResourceCodeId());
                if (Objects.nonNull(newMatCode)) {
                    isError = true;
                    isAdd = false;
                    if (StringUtils.isNotEmpty(newMatCode.getResourceCode())) {
                        //限制应急资源编码不能重复 Limit emergency resource coding cannot be repeated!
                        //记录重复的resourceCode
                        buffer.append(newMatCode.getResourceCode()+"; ");
                    }
                }
                //插入数据
                if(isAdd){
                    matCodes.add(matCode);
                }
            }

            if(isError){
                return new BaseResult("emergency.common.matCodeImportError", buffer.toString(),  false);
            }
            if (Objects.nonNull(matCodes) && matCodes.size() != 0) {
                this.addMatCodeBatch(matCodes);
            }
        } catch (Exception e) {
            logger.error("应急资源编码信息导入失败:" + e.getMessage(), e);
            ObjectUtil.throwErrorWithMsgCode("MAT_CODE_IMPORT_ERROR");

        }
        return new BaseResult(null, true);
    }

    /**
     * 批量新增
     *
     * @param matCodes
     */
    @Transactional(rollbackFor = Exception.class, value = "dataSourceManager")
    public void addMatCodeBatch(List<MatCode> matCodes) {
        for (MatCode matCode : matCodes) {
            //应急资源编码必填
            ObjectUtil.isEmptyWithMsgCode(matCode.getResourceCode(), "THE_RESOURCE_CODE_IS_NOT_NULL");
            //资源名称必填
            ObjectUtil.isEmptyWithMsgCode(matCode.getResourceName(), "RESOURCE_NAME_IS_NOT_NULL");
            //应急资源类别必填
            ObjectUtil.isEmptyWithMsgCode(matCode.getResourceType(), "RESOURCE_TYPE_IS_NOT_NULL");
            //规格型号必填
            ObjectUtil.isEmptyWithMsgCode(matCode.getSpecificationModel(), "SPECIFICATION_MODEL_IS_NOT_NULL");
            //保存操作
            matCode.setResourceCodeId(UuidUtil.getUuid());
            ControllerUtil.addMustFields(matCode);
        }
        try {
            //批量保存
            matCodeDao.addMatCodeBatch(matCodes);
        } catch (Exception e) {
            logger.error("保存应急资源编码信息失败:" + e.getMessage(), e);
            ObjectUtil.throwErrorWithMsgCode("MAT_CODE_ADD_ERROR");
        }
    }

}
