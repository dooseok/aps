<%
  '=======================================================================================================================================
  'LICENSE:
  'By Matt Parnell <parnellm@evangel.edu>
  '(C) 2013 Evangel University, all rights reserved
  'By using this software, you agree to share all improvments and changes, and leave this licensing section intact
  '=======================================================================================================================================
  'EXAMPLE:
  set oCMDInc = CreateObject("ADODB.Command")
  oCMDInc.ActiveConnection = oConnPAdmin 
  '
  strVars = "DECLARE @X AS INT; SET @X = ?; " 
  strCols = "ID,Foo,Bar,Baz"_
  strSQLBody = " FROM Table" '(joins can go here too)
  strSQLWhere = "WHERE 1=1 AND Cats = true "
  strSearchCols = "Bar,Baz" 'not required, just optional if you want certain columns specifically searchable.
  oCMDInc.Parameters.Append oCMDInc.CreateParameter("@X", adInteger, adParamInput, 8, 3382)
  
  outputDatatableJSON strVars, strCols, strSQLBody, strSQLWhere, strSearchCols, oCMDInc
  '=======================================================================================================================================
  'Outputs paginated, dynamic JSON for use by jQuery datatables
  'Notes: 
  '   if there isn't a WHERE clause, use WHERE 1=1 to avoid errors
  '   The first column in the selection list needs to be a valid column for ordering using ROW_NUMBER() OVER([first column])
   
    public sub outputDatatableJSON(strVars, strCols, strSQLBody, strSQLWhere, strSearchCols, ByRef oCMDInc)
    if len(strCols) > 0 and len(strSQLBody) > 0 then
      'Datatables specific variables to query, echo, search, paginate, etc
      if not isempty(request("sEcho")) then sEcho = Cint(Request("sEcho")) else sEcho = 0
      if not isempty(request("iDisplayLength")) then iDisplayLength = Cint(Request("iDisplayLength")) else iDisplayLength = 10
      if not isempty(request("iDisplayStart")) then iDisplayStart = Cint(Request("iDisplayStart")) else iDisplayStart = 0
      if not isempty(request("sSearch")) then sSearch = Request("sSearch") else sSearch = ""
      iTotalRecords = 0
      iTotalDisplayRecords = 0
               
      'get the number of columns (minus the num pagination column and the iTotalRows columns) we're selecting
      intNumCols = UBound(split(strCols, ","))+1
       
      'Needed for the ROW_NUMER OVER() call
      strFirstCol = split(strCols, ",")(0) 
       
      'make sure we output as JSON
      response.ContentType = "application/json"
       
      'First, get iTotalRecords into a variable
      strSQLInc = strVars & " SELECT COUNT(*) AS [iTotalDisplayRecords] " & strSQLBody & strWhere & ";"
       
      'Get our totals and proceed if we have results
      oCMDInc.CommandText = strSQLInc
      set orsInc = oCMDInc.execute
       
      if not orsInc.eof then 
        iTotalDisplayRecords = orsInc("iTotalDisplayRecords")
      else 
        iTotalRecords = 0
        iTotalDisplayRecords = 0
      end if
       
      if iTotalDisplayRecords > 0 then
       'Append the search terms
        if sSearch <> "" then
          strVars = strVars & " DECLARE @SSEARCH AS VARCHAR(200); SET @SSEARCH = ?; "
          addParam sSearch, true, oCMDInc
          arrSearch = split(strSearchCols, ",")
          i = 0
          for each searchCol in arrSearch
            if i = 0 then strClause = "AND" else strClause = "OR"
            strSQLWhere = strSQLWhere & strClause & " (" & searchCol & " LIKE @SSEARCH) "
            i = i + 1
          next
        end if  
 
        'get iTotalRecords
        strSQLInc = strVars & "SELECT COUNT(*) AS [iTotalRecords] " & strSQLBody & strSQLWhere
        oCMDInc.CommandText = strSQLInc
        set orsInc = oCMDInc.execute
        if not orsInc.eof then iTotalRecords = orsInc("iTotalRecords")
  
        'Ordering
        strOrderBy = ""
        for k = 0 to (intNumCols)
          if Request("bSortable_" & k) = "true" then
            intSortCol = int(Request("iSortCol_" & k) + 1)
            strSortDir = Request("sSortDir_" & k)
            if not isempty(strSortDir) then strOrderBy = strOrderBy & intSortCol & " " & strSortDir
          end if
        next
        if strOrderBy <> "" then strOrderBy = " ORDER BY " & strOrderBy
 
        'Select the data and iTotalRecords
        strSQLInc = strVars & "Set NOCOUNT ON; SELECT TOP " & iDisplayLength & " * FROM ( "_
          & "SELECT " & strCols & ", ROW_NUMBER() OVER (ORDER BY "&strFirstCol&") AS [num] " & strSQLBody & strSQLWhere
 
        'Append the limitation for our pagination
        strSQLInc = strSQLInc & ") T WHERE num > "& iDisplayStart & strOrderBy  
 
        'Run the query
        oCMDInc.CommandText = strSQLInc
        'showcmdmerged(ocmdinc)
        set orsInc = oCMDInc.execute
        clearCmd oCMDInc
       
        'If it isn't empty, get the total records, and pass the data into an array
        if not orsInc.eof then 
          arrResults = orsInc.GetRows()
        else 
          arrResults = ""
        end if
      end if
       
      'output the JSON 
%>{"sEcho":<%=sEcho%>,"iTotalDisplayRecords":<%=iTotalRecords%>,"iTotalRecords":<%=iTotalDisplayRecords%>,"aaData":[<%
      if iTotalDisplayRecords > 0 and isarray(arrResults) then
        For i = LBound(arrResults, 2) To UBound(arrResults, 2)
          if i > 0 then response.write ","
          %>[<%
            strThisDataPoint = ""
            for z = 0 to (intNumCols -1)
              strThisDataPoint = toUnicode(arrResults(z, i))
              if z > 0 then response.write ","
              response.write """" & strThisDataPoint & """"
              next
            %>]<%
        Next
        Erase arrResults
      end if
    else
%>{"sEcho":<%=sEcho%>,"iTotalRecords":<%=iTotalRecords%>,"iTotalDisplayRecords":<%=iTotalDisplayRecords%>,"aaData":[]<%
    end if
%>]}<%
    end sub
         
    'Adds a parameter based on it's type
    sub addParam(value, ByRef bIsSearch, ByRef oCMDInc) 'we pass byref because we don't need copies
      limit = 52 'sane value for the parameters, since user entry
      if bIsSearch then theValue = "%" & value & "%" else theValue = value
      if not isnull(theValue) and not isempty(theValue) then
          if not bIsSearch and isnumeric(theValue) then
            oCMDInc.Parameters.Append oCMDInc.CreateParameter("@SearchValue", adInteger, adParamInput, limit, theValue)
          else
            oCMDInc.Parameters.Append oCMDInc.CreateParameter("@SearchValue", adVarChar, adParamInput, limit, theValue)
          end if
        end if
  end sub
   
  'Formats the str value to a javascript friendly unicode string/value
  'Borrowed from http://www.tek-tips.com/faqs.cfm?fid=6410
  function toUnicode(str)
    dim x
    dim uStr
    dim uChr
    dim uChrCode
    iLen = len(str)
    uStr = ""
    if not isnull(iLen) then
        for x = 1 to iLen
        uChr = mid(str,x,1)
        uChrCode = asc(uChr)
        if uChrCode = 8 then ' backspace
            uChr = "\b"
        elseif uChrCode = 9 then ' tab
          uChr = "\t"
        elseif uChrCode = 10 then ' line feed
          uChr = "\n"
        elseif uChrCode = 12 then ' formfeed
          uChr = "\f"
        elseif uChrCode = 13 then ' carriage return
          uChr = "\r"
        elseif uChrCode = 34 then ' quote
          uChr = "\"""
        elseif uChrCode = 39 then ' apostrophe
          uChr = "&#39;" 'this should fix maaaany issues
        elseif uChrCode = 92 then ' backslash
          uChr = "\\"
        elseif uChrCode < 32 or uChrCode > 127 then ' non-ascii characters
          uChr = "\u" & right("0000" & CStr(uChrCode),4)
        end if
        uStr = uStr & uChr
        next
    end if
    toUnicode = uStr
  end function
%>