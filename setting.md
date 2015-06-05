# 服务器设置 #

  * 编辑manager.conf

添加如下3行代码到文件末尾:
<br>
<b><code>[</code>yourname<code>]</code></b>           ;把 yourname 修改为你的登录名<br>
<b>secret = password</b>   ;把password改成你的密码<br>
<b>read  = call</b>

<ul><li>查看 <code>[</code>general<code>]</code> 中 port和 bindaddr 的值</li></ul>

port的值默认是5038,如非必要无需更改<br>
bindaddr默认是0.0.0.0,如非必要无需更改<br>

<h1>软件设置</h1>
<table><thead><th>字段</th><th>设置说明</th></thead><tbody>
<tr><td>服务器地址</td><td>填入你服务器的域名或ip</td></tr>
<tr><td>端口</td><td>填入<code>[</code>general<code>]</code>中的port值</td></tr>
<tr><td>登录名</td><td>填入在manager.conf中设置的登录名</td></tr>
<tr><td>密码</td><td>填入在manager.conf中的设置的密码(secret)</td></tr>
<tr><td>弹屏地址</td><td>填入要弹屏的地址,用%s表示来电号码,如果不写%s的话,将自动加到末尾</td></tr>
<tr><td>分机</td><td>填写要弹屏的分机,如果所有分机都弹屏,则这一栏清空</td></tr>
<tr><td>气泡提示</td><td>选中后,来电时不仅弹屏,还会在托盘上弹出个小气泡</td></tr>
<tr><td>打开监控框</td><td>不必多说</td></tr>