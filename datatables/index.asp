<html>
<head>
<link rel="stylesheet" type="text/css" href="//cdn.datatables.net/1.10.15/css/jquery.dataTables.css">
  
<script type="text/javascript" charset="utf8" src="//cdn.datatables.net/1.10.15/js/jquery.dataTables.js"></script>
</head>
<body>
	
	
<table id="example" class="display" cellspacing="0" width="100%">
        <thead>
            <tr>
                <th>Name</th>
                <th>Position</th>
                <th>Office</th>
                <th>Extn.</th>
                <th>Start date</th>
                <th>Salary</th>
            </tr>
        </thead>
 
        <tfoot>
            <tr>
                <th>Name</th>
                <th>Position</th>
                <th>Office</th>
                <th>Extn.</th>
                <th>Start date</th>
                <th>Salary</th>
            </tr>
        </tfoot>
    </table>
    
    
    <script data-require="jquery@1.10.1" data-semver="1.10.1" src="http://code.jquery.com/jquery-1.10.1.min.js"></script>
        <script src="//cdn.datatables.net/1.10.1/js/jquery.dataTables.js"></script>
        <script type="text/javascript" data-require="angular.js@1.2.15" data-semver="1.2.15" src="http://code.angularjs.org/1.2.15/angular.js"></script>
        <script type="text/javascript" src="https://rawgithub.com/l-lin/angular-datatables/dev/dist/angular-datatables.min.js"></script>
        <script>
        $(document).ready(function() {
            $('#example').dataTable( {
                "bServerSide": true,
                "fnServerData": function (sSource, aoData, fnCallback) {
                    var myData = JSON.stringify(aoData);
                    $.ajax({
                        "dataType": 'json',
                        "contentType" : "application/json;charset=utf-8",
                        "type": "GET",
                        "url": "http://localhost:3000/ajax/phenotypes/withOrg/like/datatables/",
                        "data": myData,
                        "success": fnCallback,
                        "error": function () {
                            alert('have some problem');
                        }
                    });
                }                   
            } );
        } );
        </script>
</body>   