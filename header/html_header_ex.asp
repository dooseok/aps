<link rel="stylesheet" href="/library/flatpickr/flatpickr.min.css">
<link rel="stylesheet" type="text/css" href="/library/flatpickr/material_blue.css">
<link rel="stylesheet" href="/library/flatpickr/light.min.css">
<!--[if IE 9]>
<link rel="stylesheet" type="text/css" href="/library/flatpickr/ie.css">
<![endif]-->
<script type="text/javascript" src="/library/flatpickr/flatpickr.js"></script>
<!--<script type="text/javascript" src="/library/flatpickr/ko.js"></script>-->
<script src="/library/flatpickr/shortcut-buttons-flatpickr.min.js"></script>

<%
dim strFlatPickrArg
strFlatPickrArg = ""
strFlatPickrArg = strFlatPickrArg & "plugins: [ShortcutButtonsPlugin({"
strFlatPickrArg = strFlatPickrArg & "	button: [{label:""Today""}, {label: ""Clear""}],"
strFlatPickrArg = strFlatPickrArg & "		onClick: (index, fp) => {"
strFlatPickrArg = strFlatPickrArg & "			switch (index) {"
strFlatPickrArg = strFlatPickrArg & "				case 0:fp.setDate(new Date());break;"
strFlatPickrArg = strFlatPickrArg & "				case 1:fp.clear();break;"
strFlatPickrArg = strFlatPickrArg & "			}"
strFlatPickrArg = strFlatPickrArg & "		}})]"
%>