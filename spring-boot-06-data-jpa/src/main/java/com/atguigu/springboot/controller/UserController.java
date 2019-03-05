package com.atguigu.springboot.controller;

import com.atguigu.springboot.entity.User;
import com.atguigu.springboot.repository.UserRepository;
import com.atguigu.springboot.utils.ExportExcel;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.management.Query;
import javax.persistence.EntityManager;
import javax.servlet.http.HttpServletResponse;
import java.lang.reflect.InvocationTargetException;
import java.util.List;

@Controller
@RequestMapping("/")
public class UserController {

    @Autowired
    UserRepository userRepository;

    @Autowired
    EntityManager em;
    @GetMapping("/user/{id}")
    public User getUser(@PathVariable("id") Integer id){
        User user = userRepository.findOne(id);
        return user;
    }
    @GetMapping(value = "/all")
    public List<User> getList(){
      return  userRepository.findAll();

    }
 /*   @GetMapping(value = "/export")
    public void exportExport(HttpServletResponse response)throws InvocationTargetException {

        String[] headers = {"序号","姓名","邮箱"};
        String[] fieldNames = {"id","lastName","email"};
      List<User> userList = userRepository.findAll();
        ExportExcel.export(response, headers, fieldNames, "测试表格", userList);
    }*/
 @GetMapping(value = "/exportBaobei")
 public void exportBaobei(HttpServletResponse response)throws InvocationTargetException {

     String sql = "SELECT QuotationType AS 订单类型, a.DocumentTitle AS 项目名称, c.DistributorName 总代,d.GrossMargin 毛利率, b.RealExecutorName 审批人,c.OverSeasSalesCountry 项目销售区域, b.WFRemark 节点名称,b.OperateDateTime 审批时间 FROM dbo.Form_DocumentBase a\n" +
             "INNER JOIN dbo.Approve_WorklistLog b ON a.DocumentID=b.DocumentID\n" +
             "INNER JOIN dbo.Form_DocumentMain c ON a.DocumentID=c.DocumentID\n" +
             "INNER JOIN dbo.Form_DocumentDetails d ON a.DocumentID=d.DocumentID\n" +
             "WHERE a.QuotationType='国际业务报备' AND MONTH(a.LastModifyDateTime) IN(12) AND YEAR(a.LastModifyDateTime)=2018\n" +
             "ORDER BY a.DocumentTitle,b.OperateDateTime";
     javax.persistence.Query query = em.createNativeQuery(sql,User.class);
     String[] headers = {"订单类型","项目名称","总代","毛利率","审批人","项目销售区域","节点名称","审批时间"};
     String[] fieldNames = {"QuotationType ","DocumentTitle ","DistributorName","GrossMargin","RealExecutorName","OverSeasSalesCountry","WFRemark","OperateDateTime"};
     List<User> userList = query.getResultList();
     ExportExcel.export(response, headers, fieldNames, "国际业务报备", userList);
 }
    @GetMapping(value = "/exportXiadan")
    public void exportXiadan(HttpServletResponse response)throws InvocationTargetException {

        String sql = "SELECT QuotationType AS 订单类型, a.DocumentTitle AS 项目名称, c.DistributorName 总代,d.GrossMargin 毛利率,b.DealerName 审批人,b.DealContent 节点名称,b.DealTime 审批时间 FROM dbo.Form_DocumentBase a\n" +
                "INNER JOIN dbo.wf_Lane b ON a.ApplicationID=b.ApplicationID\n" +
                "INNER JOIN dbo.Form_DocumentMain c ON a.DocumentID=c.DocumentID\n" +
                "INNER JOIN dbo.Form_DocumentDetails d ON a.DocumentID=d.DocumentID\n" +
                "WHERE a.QuotationType='国际业务下单' AND MONTH(a.LastModifyDateTime) IN(12) AND YEAR(a.LastModifyDateTime)=2018\n" +
                "ORDER BY a.DocumentTitle,b.DealTime";
        javax.persistence.Query query = em.createNativeQuery(sql,User.class);
        String[] headers = {"订单类型","项目名称","总代","毛利率","审批人","项目销售区域","节点名称","审批时间"};
        String[] fieldNames = {"QuotationType ","DocumentTitle ","DistributorName","GrossMargin","RealExecutorName","OverSeasSalesCountry","WFRemark","OperateDateTime"};
        List<User> userList = query.getResultList();
        ExportExcel.export(response, headers, fieldNames, "国际业务下单", userList);
    }

   /* @GetMapping("/user")
    public User insertUser(User user){
        User save = userRepository.save(user);
        return save;
    }
    @RequestMapping("/show")
    public  String  show(){
        return "/show";
    }*/
}
