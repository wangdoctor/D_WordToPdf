<%@ Page Title="" Language="C#" MasterPageFile="~/Site1.Master" AutoEventWireup="true" CodeBehind="WebForm2.aspx.cs" Inherits="WebApplication2.WebForm2" %>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <h1>代码均为后台hardcode，测试功能如下，</h1>
    <h6>1.替换数据</h6>
    <p>${姓名}=>BlockChain</p>
    <p>${合同编号}=>0000111122223333</p>
    <h6>2.插入表格</h6>
    <p>
        位置在：<br />
        "合同签署所有操作均有系统、应用级日志，以天为单位将日志文件打包存证于存证平台，以此确保所有数据操作行为的可追溯，提升电子证据的证据效力。相关日志文件如下："之后的位置
    </p>
    <p>下载测试word<a download="test.docx" href="./test.docx">test.docx模板</a></p>
    <p>下载生成的pdf，执行成功之后再下载,见顶部生成结果</p>
    <span style="font-size: 18px;">
        <div>
            <input id="File1" type="file" runat="server" />
            <asp:Button ID="btnConvert" runat="server" Text="转换" OnClick="btnConvert_Click" />
        </div>
    </span>

</asp:Content>