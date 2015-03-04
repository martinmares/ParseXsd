@SET FILENAME=oreilly
@SET DOCDIR=doc

mkdir %DOCDIR%

parsexsd.rb ^
  --xsd="%FILENAME%.xsd" ^
  --xlsx="%DOCDIR%/%FILENAME%.xlsx" ^
  --request-end-with=Request ^
  --response-end-with=Response ^
  --auto-filter ^
  --frozen ^
  --indent ^
  --columns="name,schematype,type,length,multi,enum,kind,desc,mandatory,complex,simple,minoccurs,maxoccurs,nill"
