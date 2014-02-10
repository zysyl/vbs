Program PCI32;

{$apptype console}
{$r icon.res}

{$N+}
{$E+}

uses windows, sysutils, gwiopm;


{$I classes.pas}



{
  This code is Written by Craig Hart in 1996-2005. It is released as freeware;
  please use and modify at will. No gurarantees are made or implied.

  This is a derivitive work, based solely on my own PCI program for DOS.

  Please read the accompaning documentation README.RTF for all the info
  relating to this program!

}


const
  revision      : string[5]='1.6';
  busnames	: array[0..6] of string[12] =
  ('Unknown','PCI','AGP','PCI Express','CardBus','PCI-X','Empty Bus');			{ 'unknown' & 'empty bus' are never displayed onscreen }

type treedata =
  record
    b,
    d,
    f		: byte;
    vid,
    did		: word;
    cl,
    su,
    pi,
    secbus,
    subbus	: byte;
  end;

var
  olb,
  lw,
  textattr,
  PCI_hibus,
  deviceid,
  func,
  info,
  lb,
  bus,
  disp,
  cap_ptr,
  max,
  ss2c,
  ss2d,
  ss2e,
  ss2f		: byte;
           
  novpddata,
  notlast,
  debug,
  showtree,
  pdatareg,
  apic,
  found,
  businfo,
  userev,
  summary,
  bogusid,
  genssid,
  dumpregs,
  failed,
  first,
  installermode	: boolean;

  infotbl	: array[0..$ff] of byte;
  vpd		: array[0..255] of byte;
  bustype	: array[0..255] of byte;			{ Identified type of each bus }
  busvalid	: array[0..255] of boolean;			{ true = bus known to exist }
  irqmap        : array[0..255] of byte;
  pdata		: array[0..15] of byte;
  pscale	: array[0..15] of byte;
  presult	: array[0..15] of byte;
  li		: array[1..4] of string[3];
  tree		: array[0..511] of treedata;			{ list of devices and empty valid busses }
  
  cbbase,
  cblimit,
  timeout,
  pmcsr,
  romsize,
  romresult,
  rom_backup	: longint;

  spl		: real;

  vp,
  upto,
  count,
  empty,
  mb,
  ml,
  conmap,
  len,
  addr,
  index,
  i,
  j,
  l,
  v		: word;

  f		: text;

  key,
  revchk,
  oemidnum,
  oemidstr,
  cmdstr,
  vstr,
  cmpstr	: string;
  
  opsys		: string[25];







function cvtb(b:byte) : byte;
begin
  if b>9 then cvtb:=b+ord('A')-10 else cvtb:=b+ord('0');
end;

function wrhexb(byt:byte): string;
begin
 wrhexb:=chr(cvtb(byt and $0f));
end;

function wrhex(byt:byte) : string;
begin
  wrhex:=chr(cvtb((byt and $f0) shr 4))+chr(cvtb(byt and $0f));
end;

function wrhexw(wor:word): string;
begin
  wrhexw:=chr(cvtb(wor shr 12))+chr(cvtb((wor shr 8) and $f))+chr(cvtb((wor shr 4) and $f))+chr(cvtb(wor and $f));
end;

function wrhexl(lng:longint) : string;
var
  i	: integer;
  t	: string[8];
begin
  t:='';
  for i:=0 to 7 do t:=chr(cvtb((lng shr (i*4)) and $f))+t;
  wrhexl:=t;
end;  

(* Make the PCI configuration status register printout pretty *)
(* Input = the string to be output *)

Procedure printstatus (s : string);
Begin
  if not first then write(', ');
  write(s);
  first:=false;
End;


function IORedirected : boolean;
var
   hStdOut: THandle;

begin
   ioredirected:=true;
   hStdOut := GetStdHandle(STD_OUTPUT_HANDLE) ;
   if GetFileType(hStdOut)=file_type_char then IORedirected:=false;
end;




Procedure TextColor(Color: byte);
Begin
  TextAttr:=(Color AND $0F) OR (TextAttr AND $F0);
  SetConsoleTextAttribute(GetStdHandle(std_output_handle),TextAttr);
End;



function lookup_hw(deviceid,func,bus:byte;index:word) : byte;
var inf:byte;

begin
  asm
    pusha
    pushad
    push es
    push ds
    push fs
    push gs

    mov ax,$8000
    mov al,bus
    shl eax,16

    mov ax,index
    and ax,00fch
    mov ah,deviceid
    shl ah,3
    or ah,func

    mov dx,0cf8h
    out dx,eax

    mov ax,index
    and ax,3
    mov bl,8
    mul bl
    mov cx,ax

    mov dx,0cfch
    in eax,dx
    shr eax,cl
    mov inf,al

    xor eax,eax
    mov dx,0cf8h
    out dx,eax

    pop gs
    pop fs
    pop ds
    pop es
    popad
    popa
  end;
  lookup_hw:=inf;
end;


{ do NOT convert to assembler - messes with variables for some reason.. ? }
procedure write_dword_hw(deviceid,func,bus:byte;index,datah,datal:word); 
begin
  asm
    pushad

    mov ax,$8000
    mov al,bus
    shl eax,16			{ eax 31..16 }
 
    mov ax,index
    and ax,00fch
    mov ah,deviceid
    shl ah,3
    or ah,func			{ eax 15..0 }
 
    mov dx,0cf8h
    out dx,eax
 
    mov ax,datah
    shl eax,16
    mov ax,datal		{ load eax with data }
    mov dx,0cfch
    out dx,eax

    xor eax,eax			{ turn off bit 31 - no more config transactions }
    mov dx,0cf8h
    out dx,eax

    popad
  end;
end;

  

procedure lookupven(silent:boolean);
begin
  reset(f);
  failed:=true;
  repeat
    readln(f,vstr);
    if length(vstr)<3 then vstr:=';   oops';
    if (vstr[1]='V') and (copy(vstr,3,4)=cmpstr) then
    begin
      textcolor(14);
      if not silent then write(copy(vstr,8,length(vstr)));
      textcolor(7);
      failed:=false;
    end;
  until eof(f) or not failed;
  if failed and not silent then
  begin
    textcolor(12);
    write('Unknown');
    textcolor(7);
  end;
end;


procedure lookupdev;
begin
  revchk:='';
  failed:=true;
  if not eof(f) then
  repeat
    readln(f,vstr);
    if length(vstr)<3 then vstr:=';   oops';      
    if (vstr[1]='D') and (copy(vstr,3,4)=cmpstr) then
    begin
      if not eof(f) then readln(f,revchk);
      if revchk[1]='R' then
      repeat
	if wrhex(infotbl[8])=copy(revchk,3,2) then vstr:='xxxxxxx'+copy(revchk,6,length(revchk));
	if not eof(f) then readln(f,revchk);
      until revchk[1]<>'R';
      textcolor(14);
      write(copy(vstr,8,length(vstr)));
      failed:=false;
      textcolor(7);
    end;
  until eof(f) or not failed or (vstr[1]='V');
  if failed then
  begin
    textcolor(12);
    write('Unknown');
    textcolor(7);
  end;
end;



procedure showinstallerinfo;
begin
  write('V:',wrhexw(infotbl[1] shl 8+infotbl[0]),' ');
  write('D:',wrhexw(infotbl[3] shl 8+infotbl[2]),' ');
  write('S:');
  if infotbl[$e] and $7f=0 then
  begin
    write(wrhexw(infotbl[$2f] shl 8+infotbl[$2e]));
    write(wrhexw(infotbl[$2d] shl 8+infotbl[$2c]),' ');
  end else write('00000000 ');
  write('B:',bus,' ');
  write('E:');
  if deviceid<10 then write('0');
  write(deviceid,' ');
  write('F:',func,' ');
  write('I:',wrhex(infotbl[$3c]),' ');
  write('N:');
  if infotbl[$3c]>0 then
  begin
    if infotbl[$3d]=0 then write('- ') else write(chr(infotbl[$3d]+64),' ');
  end else write('- ');
  write('C:',wrhex(infotbl[$b]),' ');
  write('U:',wrhex(infotbl[$a]),' ');
  write('P:',wrhex(infotbl[$9]),' ');
  write('R:',wrhex(infotbl[$8]));
  writeln;
end;




procedure docapdecode;
var
  i,
  j	: word;
  
begin
  writeln(' New Capabilities List Present:');
{type 0}     if infotbl[$e] and $7f=0 then cap_ptr:=infotbl[$34];
{type 1}     if infotbl[$e] and $7f=1 then cap_ptr:=infotbl[$34];
{type 2}     if infotbl[$e] and $7f=2 then cap_ptr:=infotbl[$14];

  if cap_ptr<>0 then
  repeat
    case infotbl[cap_ptr] of

      01 : begin
	     write('   Power Management Capability, Version ');
	     if infotbl[cap_ptr+2]and 7=3 then writeln('1.2') else
	     if infotbl[cap_ptr+2]and 7=2 then writeln('1.1') else
	     if infotbl[cap_ptr+2]and 7=1 then writeln('1.0') else
	     begin
	       textcolor(12);
	       writeln('Unknown: Code ',wrhexb(infotbl[cap_ptr+2] and 7),'h');
	       textcolor(7);
	     end;


{ list supported low-power states; D3 and D0 are always supported}
	     if infotbl[cap_ptr+3] and 2=2 then writeln('     Supports low power State D1');
	     if infotbl[cap_ptr+3] and 4=4 then writeln('     Supports low power State D2');
	     if infotbl[cap_ptr+3] and 6=0 then writeln('     Does not support low power State D1 or D2');

{list PME# generation capabilities}	     
	     if (infotbl[cap_ptr+3] shr 3) = 0 then writeln('     Does not support PME# signalling') else
	     begin
	       write('     Supports PME# signalling from mode(s) ');
	       first:=true;
	       if (infotbl[cap_ptr+3] shr 3) and 1=1 then printstatus('D0');
	       if (infotbl[cap_ptr+3] shr 3) and 2=2 then printstatus('D1');
	       if (infotbl[cap_ptr+3] shr 3) and 4=4 then printstatus('D2');
	       if (infotbl[cap_ptr+3] shr 3) and 8=8 then printstatus('D3hot');
	       if (infotbl[cap_ptr+3] shr 3) and 16=16 then printstatus('D3cold');
	       writeln;
	       
	       write('     PME# signalling is currently ');
	       if infotbl[cap_ptr+5] and 1=1 then writeln('enabled') else writeln('disabled');
	      end;
	     

	     write('     Current Power State : D');
	     case infotbl[cap_ptr+4] and 3 of
	       0 : writeln('0 (Device operational, no power saving)');
	       1 : writeln('1 (Device idle, minimum power saving)');
	       2 : writeln('2 (Device CLK stopped, medium power saving)');
	       3 : writeln('3hot (Device off: no power to device, maximum power saving)');
	     end;


{ try to read the power levels, but only if card not asleep now. }
	     if infotbl[cap_ptr+4] and 3=0 then 
	     begin
{ backup registers we are going to mess with }     
	     pmcsr:=infotbl[cap_ptr+7] shl 24 + infotbl[cap_ptr+6] shl 16 + infotbl[cap_ptr+5] shl 8 + infotbl[cap_ptr+4];

	     for i:=0 to 15 do
	     begin
	       write_dword_hw(deviceid,func,bus,cap_ptr+4,0000,(((infotbl[$e1] shl 8) + infotbl[$e0]) and $0103) + (i shl 9));
	       presult[i]:=(lookup_hw(deviceid,func,bus,cap_ptr+5) shr 1) and $f;
	       pdata[i]:=lookup_hw(deviceid,func,bus,cap_ptr+7);
	       pscale[i]:=(lookup_hw(deviceid,func,bus,cap_ptr+5) shr 5) and 3;
	     end;

{ restore original values to registers regardless }
	     write_dword_hw(deviceid,func,bus,cap_ptr+4,pmcsr shr 16,pmcsr and $ffff);

{ is the reg valid ? }
	     pdatareg:=false;
	     for i:=0 to 15 do if (pdata[i]<>0) or (pscale[i]<>0) then pdatareg:=true;
	     
	     if pdatareg then
	     begin
	       writeln('     Power Data Registers Information:');

{ debug: for i:=0 to 15 do writeln('Index: ',i,', Data: ',pdata[i],', Scale: ',pscale[i]);}



	       if infotbl[$e] and $80=0 then max:=7 else max:=8;
	       for i:=0 to max do
	       begin
	         case i of
	           0..3 : write('      D',i,' Power Consumed: ');
	           4..7 : write('      D',i-4,' Power Dissipated: ');
	           8 :  write('      Common logic Power Consumed: ');
	         end;
	         case pscale[i] of
	           0 : write('???');
	           1 : write(pdata[i]*100);
	           2 : write(pdata[i]*10);
	           3 : write(pdata[i]);
	         end;
	         writeln('mW');
	       end;
	     end else
	     begin
	       if (infotbl[cap_ptr+3] shr 3) and 16=16 then 
	       begin
	         write('     3.3v AUX Current required : ');
	         case ((infotbl[cap_ptr+3] and 1) shl 2) + (infotbl[cap_ptr+2] shr 6) of
	           0 : writeln('0mA (Self powered)');
	           1 : writeln('55mA');
	           2 : writeln('100mA');
	           3 : writeln('160mA');
	           4 : writeln('220mA');
	           5 : writeln('270mA');
	           6 : writeln('320mA');
	           7 : writeln('375mA');
	         end;
	       end;
	     end;
	     end;
	   end;


      02 : begin
	     write('   AGP Capability, Version ');
	     write(infotbl[cap_ptr+2] shr 4,'.',infotbl[cap_ptr+2] and $0f,' ');
	     if (infotbl[cap_ptr+2] shr 4)=1 then writeln('(AGP 1x and/or 2x support)');
	     if (infotbl[cap_ptr+2] shr 4)=2 then writeln('(AGP 4x and below support)');
	     if (infotbl[cap_ptr+2] shr 4)=3 then
	     begin
	       write('(AGP 8x and 4x');
	       if (infotbl[cap_ptr+2] and $0f)<5 then write(', core register');
	       if ((infotbl[cap_ptr+2] and $0f)>4) and
		 ((infotbl[cap_ptr+2] and $0f)<10) then write(', appendix register');
	       writeln(' support)');
	     end;


	     write('     AGP Mode Enabled : ');
	     if infotbl[cap_ptr+9] and 1=1 then
	     begin
               textcolor(10);
	       writeln('Yes');
               textcolor(7);
	     end else
	     begin
               textcolor(12);
               writeln('No');
               textcolor(7);
	     end;


	     write('     AGP Speed(s) Supported :');
	     if infotbl[cap_ptr+4] and 8=8 then
	     begin
	       if infotbl[cap_ptr+4] and 1=1 then write(' 4x');
	       if infotbl[cap_ptr+4] and 2=2 then write(' 8x');
	       if infotbl[cap_ptr+4] and 7>3 then write(' Unknown Speed Reported (',wrhex(infotbl[cap_ptr+4] and 7),'h)!!');
	     end else
	     begin
	       if infotbl[cap_ptr+4] and 1=1 then write(' 1x');
	       if infotbl[cap_ptr+4] and 2=2 then write(' 2x');
	       if infotbl[cap_ptr+4] and 4=4 then write(' 4x');
	       if infotbl[cap_ptr+4] and 7=0 then
	       begin
                 textcolor(12);
		 write(' None!!');
                 textcolor(11);
		 write(' (Assume Only 1x Support)');
                 textcolor(7);
	       end;
	     end;

	     write(', Currently :');
	     if infotbl[cap_ptr+4] and 8=8 then
	     begin
	       if infotbl[cap_ptr+8] and 7=1 then write(' 4x');
	       if infotbl[cap_ptr+8] and 7=2 then write(' 8x');
	       if infotbl[cap_ptr+8] and 7>2 then write(' Unknown Speed Reported (',wrhex(infotbl[cap_ptr+8] and 7),'h)!!');
	       if infotbl[cap_ptr+8] and 7=0 then write(' None Selected');
	     end else
	     begin
	       if infotbl[cap_ptr+8] and 7=1 then write(' 1x');
	       if infotbl[cap_ptr+8] and 7=2 then write(' 2x');
	       if infotbl[cap_ptr+8] and 7=4 then write(' 4x');
	       if infotbl[cap_ptr+8] and 7=0 then write(' None Selected');
	     end;
	     writeln;




	     write('     FW Transfers Supported : ');
	     if infotbl[cap_ptr+4] and $10=$10 then write('Yes') else write('No');
	     write(', Enabled : ');
	     if infotbl[cap_ptr+8] and $10=$10 then writeln('Yes') else writeln('No');

	     write('     >4Gb Address Space Supported : ');
	     if infotbl[cap_ptr+4] and $20=$20 then write('Yes') else write('No');
	     write(', Enabled : ');
	     if infotbl[cap_ptr+8] and $20=$20 then writeln('Yes') else writeln('No');

	     write('     Sideband Addressing Supported : ');
	     if infotbl[cap_ptr+5] and 2=2 then write('Yes') else write('No');
	     write(', Enabled : ');
	     if infotbl[cap_ptr+9] and 2=2 then writeln('Yes') else writeln('No');
	    

{ if v3.? reported, see if v3.0 mode is on }
	     if (infotbl[cap_ptr+2] shr 4)=3 then
	     begin
	       write('     AGP v3.0 Operation Mode Available : ');
	       if infotbl[cap_ptr+4] and 8=8 then
	       begin
	         write('Yes, Enabled : ');
	         if infotbl[cap_ptr+9] and 1=1 then writeln('Yes') else writeln('No');
	        end else writeln('No');
	     end;
	     

{ isosynch only in AGP v3.0 mode }
	     if infotbl[cap_ptr+4] and 8=8 then
	     begin
	       write('     Isosynchronous Transactions Supported : ');
	       if infotbl[cap_ptr+6] and 2=2 then writeln('Yes') else writeln('No');
	     end;


	     write('     Maximum Command Queue Length : ',infotbl[cap_ptr+7]+1,' byte');
	     if infotbl[cap_ptr+7]>0 then write('s');
	     write(', Current Queue Length : ',infotbl[cap_ptr+11]+1,' byte');
	     if infotbl[cap_ptr+11]=0 then writeln else writeln('s');


	   end;





      03 : begin
	     writeln('   Vital Product Data Capability');

	     failed:=false;
	     novpddata:=false;
	     i:=0;
	     repeat

	       write_dword_hw(deviceid,func,bus,cap_ptr,(i shl 2),infotbl[cap_ptr+1] shl 8 + infotbl[cap_ptr]);
	       timeout:=0;
	       repeat
	         infotbl[cap_ptr+3]:=lookup_hw(deviceid,func,bus,cap_ptr+3);
	         inc(timeout);
	       until (infotbl[cap_ptr+3] and $80=$80) or (timeout=1000000);
	       
	       if debug then if timeout=1000000 then
	       begin
	         textcolor(12);
	         write('T');
	         textcolor(7);
	       end else write('*');
	       
	       if timeout=1000000 then failed:=true;
	       
	       
	       vpd[i shl 2]:=lookup_hw(deviceid,func,bus,cap_ptr+4);
	       vpd[i shl 2 + 1]:=lookup_hw(deviceid,func,bus,cap_ptr+5);
	       vpd[i shl 2 + 2]:=lookup_hw(deviceid,func,bus,cap_ptr+6);
	       vpd[i shl 2 + 3]:=lookup_hw(deviceid,func,bus,cap_ptr+7);
	       
	       if debug then write(i:2,':',wrhex(vpd[i shl 2]),' ',wrhex(vpd[i shl 2 + 1]),' ',wrhex(vpd[i shl 2 + 2]),' ',wrhex(vpd[i shl 2 + 3]),' ');
	       
	       if i=0 then if (vpd[0]=vpd[1]) and (vpd[1]=vpd[2]) and (vpd[2]=vpd[3]) then
	       begin
	         failed:=true;
	         novpddata:=true;
	         writeln('     VPD Data not present');
	       end;
	       if failed then i:=64 else inc(i);
	       
	     until i=64;





	     
	     if debug then writeln;
	     
	     
	     if not novpddata then
	     begin
	     
	     if debug then writeln('!!Decode Begins');
	     
	     
	     vp:=0;
	     
	     repeat
	       if vpd[vp] and $80=0 then
	       begin
	         { small }
	         if (vpd[vp] and $7f) shr 3=$0f then
	         begin
	           { end }
	           writeln('     End Tag');
	           vp:=0;
	         end else
	         vp:=vp+(vpd[vp] and 7)+1;
	       end else
	       
	       begin
	         { big }
	         if vpd[vp] and $7f=$02 then
	         begin
	           { string }
	           write('     Identifier          ');
	           for i:=vp+3 to vp+2+(vpd[vp+2] shl 8 + vpd[vp+1]) do write(chr(vpd[i]));
	           writeln;
	           vp:=vp+(vpd[vp+2] shl 8 + vpd[vp+1])+3;
	         end;

	         if vpd[vp] and $7f=$10 then
	         begin
	           { VPD-R }
	           if debug then writeln('VPD-R Data');
	           upto:=vp+(vpd[vp+2] shl 8 + vpd[vp+1])+3;
	           vp:=vp+3;
	           repeat
	             if debug then writeln('Keyword : ',chr(vpd[vp]),chr(vpd[vp+1]),' Length : ',wrhex(vpd[vp+2]),'h : ');
	             key:=chr(vpd[vp])+chr(vpd[vp+1]);
	             if (key='FG') or (key='LC') or (key='CP') or (key='RV') or (key='RW') then found:=true else found:=false;;
	             write('     ');
	             if key='PN' then write('Part Number         ');
	             if key='EC' then write('EC Level (Version)  ');
	             if key='MN' then write('Manufacturer ID     ');
	             if key='SN' then write('Serial Number       ');
	             if key[1]='V' then write('Vendor Specific     ');
	             if key[1]='Y' then write('System Specific     ');
	             if key='RV' then write('Checksum/End        ');
	             if key='RW' then write('Remaining RW Area   ');
	             if key='CP' then write('Extended Capability ');
	             if key='PG' then write('PCI Geography       ');
	             if key='LC' then write('Location            ');
	             if key='FG' then write('Fabric Geography    ');
	             if not found then
	               for j:=1 to vpd[vp+2] do write(chr(vpd[vp+2+j])) else
	               for j:=1 to vpd[vp+2] do write(wrhex(vpd[vp+2+j]),'h ');
	             writeln;
	             vp:=vp+3+vpd[vp+2];
	           until vp=upto;
	         end;
	       
	         if vpd[vp] and $7f=$11 then
	         begin
	           { VPD-W }
	           if debug then writeln('VPD-W Data');
	           upto:=vp+(vpd[vp+2] shl 8 + vpd[vp+1])+3;
	           vp:=vp+3;
	           repeat
	             if debug then writeln('Keyword : ',chr(vpd[vp]),chr(vpd[vp+1]),' Length : ',wrhex(vpd[vp+2]),'h : ');
	             key:=chr(vpd[vp])+chr(vpd[vp+1]);
	             if (key='FG') or (key='LC') or (key='CP') or (key='RV') or (key='RW') then found:=true else found:=false;;
		     write('     ');
	             if key='PN' then write('Part Number         ');
	             if key='EC' then write('EC Level (Version)  ');
	             if key='MN' then write('Manufacturer ID     ');
	             if key='SN' then write('Serial Number       ');
	             if key[1]='V' then write('Vendor Specific     ');
	             if key[1]='Y' then write('System Specific     ');
	             if key='RV' then write('Checksum/End        ');
	             if key='RW' then write('Remaining RW Area   ');
	             if key='CP' then write('Extended Capability ');
	             if key='PG' then write('PCI Geography       ');
	             if key='LC' then write('Location            ');
	             if key='FG' then write('Fabric Geography    ');
	             if not found then
	               for j:=1 to vpd[vp+2] do write(chr(vpd[vp+2+j])) else
	               for j:=1 to vpd[vp+2] do write(wrhex(vpd[vp+2+j]),'h ');
	             writeln;
	             vp:=vp+3+vpd[vp+2];
	           until vp=upto;
	         end;
	         
	       end;
	     until vp=0;  
	     
	     end;
	   end;







      04 : begin
	     writeln('   Slot Identification Capability');
	     write('     This is ');
	     if infotbl[cap_ptr+2] and $20=0 then write('not ');
	     writeln('a parent bridge');
	     write('     Number of slots on secondary side of this bridge : ');
	     writeln(infotbl[cap_ptr+2] and $1f);
	     writeln('Chassis Number : ',infotbl[cap_ptr+3]);
	   end;


      05 : begin
	     writeln('   Message Signalled Interrupt Capability');
	     write('     MSI is ');
	     if infotbl[cap_ptr+2] and 1=1 then writeln('enabled') else writeln('disabled');

	     write('     MSI function can generate ');
	     if infotbl[cap_ptr+2] and 128=128 then write('64') else write('32');
	     writeln('-bit addresses');
	    end;


      06 : begin
	     writeln('   CompactPCI Hot-Swap Capability');
	   end;


      07 : begin
	     writeln('   PCI-X Capability');

{ type 1 }
	     if infotbl[$e] and $7f=1 then
	     begin
	       write('     Secondary AD Interface Size is ');
	       if infotbl[cap_ptr+2] and 1=1 then write('64') else write('32');
	       writeln('-bits wide');

	       write('     Secondary Bus Maximum Speed in PCI-X Mode 1 is ');
	       if infotbl[cap_ptr+2] and 2=2 then write('133') else write('66');
	       writeln('MHz');
	       
	       write('     Secondary Bus is PCI-X 266 Capable : ');
	       if infotbl[cap_ptr+3] and $40=$40 then writeln('Yes') else writeln('No');
	       write('     Secondary Bus is PCI-X 533 Capable : ');
	       if infotbl[cap_ptr+3] and $80=$80 then writeln('Yes') else writeln('No');


	       write('     Seconday Bus Current Speed : ');
	       case ((infotbl[cap_ptr+3] and 3) shl 2 + (infotbl[cap_ptr+2] and $c0) shr 6) of
		 0 : write('PCI (33MHz)');
		 1 : write('PCI-X Mode 1 (66MHz)');
		 2 : write('PCI-X Mode 1 (100MHz)');
		 3 : write('PCI-X Mode 1 (133MHz)');
		 4..7 : write('PCI-X Mode 1 (Unknown Speed!)');
		 8 : write('PCI-X 266 (Unknown Speed!)');
		 9 : write('PCI-X 266 (66MHz)');
		 $a : write('PCI-X 266 (100MHz)');
		 $b : write('PCI-X 266 (133MHz)');
		 $c : write('PCI-X 533 (Unknown Speed!)');
		 $d : write('PCI-X 533 (66MHz)');
		 $e : write('PCI-X 533 (100MHz)');
		 $f : write('PCI-X 533 (133MHz)');
	       end;
	     end;


{ type 0 and 1 }
	     write('     Primary AD Interface Size is ');
	     if infotbl[cap_ptr+6] and 1=1 then write('64') else write('32');
	     writeln('-bits Wide');

	     write('     Primary AD Bus Maximum Speed in PCI-X Mode 1 is ');
	     if infotbl[cap_ptr+6] and 2=2 then write('133') else write('66');
	     writeln('MHz');

             write('     Primary Bus is PCI-X 266 Capable : ');
             if infotbl[cap_ptr+7] and $40=$40 then writeln('Yes') else writeln('No');
	     write('     Primary Bus is PCI-X 533 Capable : ');
	     if infotbl[cap_ptr+7] and $80=$80 then writeln('Yes') else writeln('No');


	   end;


      08 : begin
             writeln('   HyperTransport Capability');
             write('     SubType : ');
             if infotbl[cap_ptr+3] and $e0=0 then writeln('Slave/Primary Interface') else
             if infotbl[cap_ptr+3] and $e0=$20 then writeln('Host/Secondary Interface') else
             case infotbl[cap_ptr+3] and $F8 of
               $40 : writeln('Switch');
               $80 : writeln('Interrupt Discovery & Configuration');
               $88 : writeln('Revision ID');
               $90 : writeln('UnitID Clumping');
               $98 : writeln('Extended Configuration Space Access');
               $A0 : writeln('Address Mapping');
               $A8 : writeln('MSI Mapping');
               $B0 : writeln('DirectRoute');
               $B8 : writeln('VCSet');
               $C0 : writeln('Retry mode');
               $C8 : writeln('x86 Encoding(Reserved)');
               else writeln('?? Unknown');
             end;

{ Slave/Pri : type 000xx}            
             if infotbl[cap_ptr+3] and $e0=0 then
             begin
               writeln('     Base UnitID  : ',infotbl[cap_ptr+2] and $1f);
               writeln('     UnitID Count : ',((infotbl[cap_ptr+3] and 3) shl 3) + (infotbl[cap_ptr+2] shr 5));
             end;

           end;


      09 : begin
	     writeln('   Vendor-Dependant Capability');
	   end;

     $0a : begin
	     writeln('   USB 2.0 EHCI Debug Port Capability');
	   end;

     $0b : begin
	     writeln('   CompactPCI Resource Control Capability');
	   end;

     $0c : begin
	     writeln('   PCI Hot-Plug Capability');
	   end;

     $0d : begin
             writeln('   Subsystem ID & Subsystem Vendor ID Capability');
             writeln('     SSVID : ',wrhexw(infotbl[cap_ptr + 5] shl 8 + infotbl[cap_ptr + 4]),'h');
             writeln('     SSID  : ',wrhexw(infotbl[cap_ptr + 7] shl 8 + infotbl[cap_ptr + 6]),'h');
           end;

     $0e : begin
	     writeln('   AGP 8x Capability');
	   end;

     $0f : begin
	     writeln('   Secure Device Capability');
	   end;

     $10 : begin
	     writeln('   PCI Express Capability, Version ',infotbl[cap_ptr+2] and $f);

{ **** 3GIO Capabilities field 2-3 }
	     writeln('     Device/Port Type : ');
	     write('       ');
	     case infotbl[cap_ptr+2] shr 4 of
	       0 : writeln('PCI Express Endpoint Device');
	       1 : writeln('Legacy PCI Express Endpoint Device');
	       4 : writeln('Root port of PCI Express Root Complex');
	       5 : writeln('Upstream port of PCI Express Switch');
	       6 : writeln('Downstream port of PCI Express Switch');
	       7 : writeln('PCI Express to PCI/PCI-X Bridge');
	       8 : writeln('PCI/PCI-X to PCI Express Bridge');
	       9 : writeln('Root Complex Integrated Endpoint Device');
	       else writeln('Unknown (',wrhex(infotbl[cap_ptr+2] shr 4),'h)!!');
	      end;

{ ports only: 4,5 (6?) }	            
	      if (infotbl[cap_ptr+2] shr 4=4) or (infotbl[cap_ptr+2] shr 4=6) then
	      begin
	        writeln('     Port Type :');
	        write('       Port is an ');
	        if infotbl[cap_ptr+3] and 1=1 then writeln('Expansion Slot') else writeln('Integrated Device');
	      end;

{ **** Device Capabilities field 4-7}	      
{ Types: 0,1,5,7 } 
	      if (infotbl[cap_ptr+2] shr 4=0) or (infotbl[cap_ptr+2] shr 4=1) or (infotbl[cap_ptr+2] shr 4=5) or (infotbl[cap_ptr+2] shr 4=7) then
	      begin
  	        writeln('     Device Capabilities :');
  	        if infotbl[cap_ptr+5] and 16=16 then writeln('       Attention Button Present on Device');
	        if infotbl[cap_ptr+5] and 32=32 then writeln('       Attention Indicator Present on Device');
	        if infotbl[cap_ptr+5] and 64=64 then writeln('       Power Indicator Present on Device');
	      end;

{ upstream ports only: 5 }
	      if infotbl[cap_ptr+2] shr 4=5 then
	      begin
	        writeln('     Upstream Port Power Data :');
  	        write('       Slot Power Limit Value : ');
		spl:=(infotbl[cap_ptr+6] shr 2) + ((infotbl[cap_ptr+7] and 3) shl 6);
  	        if ((infotbl[cap_ptr+7] and $c) shr 2)=0 then spl:=spl;
  	        if ((infotbl[cap_ptr+7] and $c) shr 2)=1 then spl:=spl*0.1;
  	        if ((infotbl[cap_ptr+7] and $c) shr 2)=2 then spl:=spl*0.01;
  	        if ((infotbl[cap_ptr+7] and $c) shr 2)=3 then spl:=spl*0.001;
  	        writeln(spl:7:3,' Watts');
  	      end;
  	       	      
{ **** Device Control Field 8-9}
  	      writeln('     Device Control :');
  	      if infotbl[cap_ptr+8] and 1=1 then writeln('       Correctable Error Reporting Enabled');
  	      if infotbl[cap_ptr+8] and 2=2 then writeln('       Non-Fatal Error Reporting Enabled');
  	      if infotbl[cap_ptr+8] and 4=4 then writeln('       Fatal Error Reporting Enabled');
  	      if infotbl[cap_ptr+8] and 8=8 then writeln('       Unsupported Request Reporting Enabled');
  	      write('       Unsupported Request Severity is ');
  	      if infotbl[cap_ptr+8] and 16=16 then writeln('Fatal') else writeln('Non-Fatal');


{ **** Device Status field 0ah-0bh}
  	      writeln('     Device Status :');
  	      if infotbl[cap_ptr+$a] and 1=1 then writeln('       Correctable Error Detected');
  	      if infotbl[cap_ptr+$a] and 2=2 then writeln('       Non-Fatal Error Detected');
  	      if infotbl[cap_ptr+$a] and 4=4 then writeln('       Fatal Error Detected');
  	      if infotbl[cap_ptr+$a] and 8=8 then writeln('       Unsupported Request Detected');
  	      if infotbl[cap_ptr+$a] and 16=16 then writeln('       AUX Power Detected');
  	      if infotbl[cap_ptr+$a] and 32=32 then writeln('       Device Reports Transactions Pending');

	      
{ **** Link Capabilities field 0ch-0fh}
	      writeln('     Link Capabilities : ');
	      write('       Maximum Link Speed : ');
	      case infotbl[cap_ptr+$c] and $f of
	        1 : writeln('2.5Gb/s');
	        else writeln('Unknown (',wrhex(infotbl[cap_ptr+$c] and $f),'h)!!');
	      end;

	      write('       Maximum Link Width : x');
	      lw:=((infotbl[cap_ptr+$d] and 3) shl 4) + ((infotbl[cap_ptr+$c] and $f0) shr 4);
	      if lw=0 then writeln('Reserved') else writeln(lw);

	      writeln('       Link Port Number   : ',infotbl[cap_ptr+$f]);

	 
{ **** Link Control 10h-11h }
  	      writeln('     Link Control :');
  	      if infotbl[cap_ptr+$10] and 4=4 then writeln('       Link is in Loopback mode');
  	      if infotbl[cap_ptr+$10] and 16=16 then writeln('       Link is Disabled');
	      if infotbl[cap_ptr+$10] and 64=64 then writeln('       Common Clock Configuration In Use') else writeln('       Asynchronous Clocking in Use');


{ **** Link Status 12h-13h } 
              writeln('     Link Status :');
              write('       Current Link Speed : ');
	      case infotbl[cap_ptr+$12] and $f of
	        1 : writeln('2.5Gb/s');
	        else writeln('Unknown (',wrhex(infotbl[cap_ptr+$12] and $f),')!!');
	      end;
	      
	      write('       Current Link Width : x');
	      lw:=((infotbl[cap_ptr+$13] and 3) shl 4) + ((infotbl[cap_ptr+$12] and $f0) shr 4);
	      if infotbl[cap_ptr+$13] and 8=8 then writeln('??') else writeln(lw);
	      if infotbl[cap_ptr+$13] and 4=4 then writeln('     Link Training Error Reported!!');
	      if infotbl[cap_ptr+$13] and 8=8 then writeln('     Link Training Currently In Progress!!');


{ **** Slot Capabilities 14h-17h, root & downstream ports that are slots only }

  	     if (((infotbl[cap_ptr+2] shr 4)=4) or ((infotbl[cap_ptr+2] shr 4)=6)) and (infotbl[cap_ptr+3] and 1=1) then
  	     begin
  	       writeln('     Slot Capabilities :');
  	       if infotbl[cap_ptr+$14] and 1=1 then writeln('       Attention Button Present');
  	       if infotbl[cap_ptr+$14] and 2=2 then writeln('       Power Controller Present');
  	       if infotbl[cap_ptr+$14] and 4=4 then writeln('       MRL Sensor Present');
  	       if infotbl[cap_ptr+$14] and 8=8 then writeln('       Attention Indicator Present');
  	       if infotbl[cap_ptr+$14] and 16=16 then writeln('       Power Indicator Present');
  	       if infotbl[cap_ptr+$14] and 32=32 then writeln('       Hot Plug Surprise is Possible');
  	       if infotbl[cap_ptr+$14] and 64=64 then writeln('       Hot Plug Capable');

 	       write('       Slot Power Limit Value : ');
	       spl:=((infotbl[cap_ptr+15] and $7f) shl 1) + (infotbl[cap_ptr+14] shr 7);
 	       if (infotbl[cap_ptr+15] shr 7) + ((infotbl[cap_ptr+16] and 1) shl 1)=0 then spl:=spl;
 	       if (infotbl[cap_ptr+15] shr 7) + ((infotbl[cap_ptr+16] and 1) shl 1)=1 then spl:=spl*0.1;
  	       if (infotbl[cap_ptr+15] shr 7) + ((infotbl[cap_ptr+16] and 1) shl 1)=2 then spl:=spl*0.01;
  	       if (infotbl[cap_ptr+15] shr 7) + ((infotbl[cap_ptr+16] and 1) shl 1)=3 then spl:=spl*0.001;
  	       writeln(spl:7:3,' Watts');

  	       writeln('       Physical slot Number ',(word(word(infotbl[cap_ptr+$17]) shl 2)+(infotbl[cap_ptr+$16] shr 6)));
  	     end;


  	   end;


     $11 : begin
	     writeln('   MSI-X Capability');
	   end;



     $12 : begin
             writeln('   SATA HBA Optional Features Present');
           end;
           

      else writeln('   Unknown Capability (Code ',wrhex(infotbl[cap_ptr]),'h)!!');
    end;
  cap_ptr:=infotbl[cap_ptr+1];
  until cap_ptr=0 else writeln('  No ''New Capabilities'' Are Currently Enabled!');
end;




procedure showallinfo;
var
  j,
  i     : integer;
  pp,
  nn	: byte;

begin
	if businfo then
	begin
	  write(' Bus ');
          textcolor(11);
	  write(bus);
          textcolor(7);
	  write(' (',busnames[bustype[bus]],'), Device Number ');
          textcolor(11);
	  write(deviceid);
          textcolor(7);
	  write(', Device Function ');
          textcolor(11);
	  writeln(func);
          textcolor(7);
	end;


	if installermode then showinstallerinfo else
	  begin

	  write(' Vendor ',wrhexw(infotbl[1] shl 8+infotbl[0]),'h ');
	  cmpstr:=wrhexw(infotbl[1] shl 8+infotbl[0]);
	  lookupven(false);
	  writeln;

	  write(' Device ',wrhexw(infotbl[3] shl 8+infotbl[2]),'h ');
	  cmpstr:=wrhexw(infotbl[3] shl 8+infotbl[2]);
	  lookupdev;
	  writeln;


	  if not summary then
	  begin
{ command register }
	    write(' Command ',wrhexw(infotbl[5] shl 8+infotbl[4]),'h');
	    first:=true;
	    write(' (');
	    if infotbl[4] and 1=1 then printstatus('I/O Access');
	    if infotbl[4] and 2=2 then printstatus('Memory Access');
	    if infotbl[4] and 3=0 then printstatus('Bus Access Disabled!!');
	    if infotbl[4] and 4=4 then printstatus('BusMaster');
	    if infotbl[4] and 8=8 then printstatus('Special Cycles');
	    if infotbl[4] and 16=16 then printstatus('MemWrite+Invalidate');
	    if infotbl[4] and 32=32 then printstatus('VGA Palette Snoop');
	    if infotbl[4] and 64=64 then printstatus('Parity Error Response');
	    if infotbl[4] and 128=128 then printstatus('Wait Cycles');
	    if infotbl[5] and 1=1 then printstatus('System Errors');
	    if infotbl[5] and 2=2 then printstatus('Back-To-Back Transactions');
	    if infotbl[5] and 4=4 then printstatus('Interrupt Disable');
	    writeln(')');

{ status register }
	    write(' Status ',wrhexw(infotbl[7] shl 8+infotbl[6]),'h');
	    if (infotbl[6]<>0) or (infotbl[7]<>0) then
	    begin
	      first:=true;
	      write(' (');
	      if infotbl[6] and 8=8 then printstatus('Signalled Interrupt');
	      if infotbl[6] and 16=16 then printstatus('Has Capabilities List');
	      if infotbl[6] and 32=32 then printstatus('Supports 66MHz');
	      if infotbl[6] and 64=64 then printstatus('Has UDF');
	      if infotbl[6] and 128=128 then printstatus('Supports Back-To-Back Trans.');

	      if infotbl[7] and 1=1 then printstatus('Data parity Error Detected');
	      if infotbl[7] and 8=8 then printstatus('Signalled Target Abort');
	      if infotbl[7] and 16=16 then printstatus('Received Target Abort');
	      if infotbl[7] and 32=32 then printstatus('Received Master Abort');
	      if infotbl[7] and 64=64 then printstatus('Signalled System Error');
	      if infotbl[7] and 128=128 then printstatus('Detected Parity Error');

	      case ((infotbl[7] and 6) shr 1) of
		0 : printstatus('Fast Timing');
		1 : printstatus('Medium Timing');
		2 : printstatus('Slow Timing');
		3 : printstatus('Unknown Timing');
	      end;
	      write(')');
	    end;
	    writeln;

{ misc. general registers }
	    write(' Revision ',wrhex(infotbl[8]),'h');
	    write(', Header Type ',wrhex(infotbl[$e]),'h');
	    writeln(', Bus Latency Timer ',wrhex(infotbl[$d]),'h');

{ header type 0 only : display latency and grant figures }
	   if infotbl[$e] and $7f=0 then
	   begin
	     if (infotbl[$3e]<>0) or (infotbl[$3f]<>0) then
	       writeln(' Minimum Bus Grant ',wrhex(infotbl[$3e]),'h, Maximum Bus Latency ',wrhex(infotbl[$3f]),'h');
	   end;

{ self test }
	    write(' Self test ',wrhex(infotbl[$f]),'h (Self test ');
	    if infotbl[$f] and $80=0 then write('not ');
	    write('supported');

	    if infotbl[$f] and $80=$80 then
	    begin
	      write(': Completion code ',wrhexb(infotbl[$f] and $f),'h - ');
	      if infotbl[$f] and $f=0 then
	      begin
                textcolor(10);
		write('OK');
                textcolor(7);
	      end else
	      begin
                textcolor(12);
		write('Failed!!');
                textcolor(7);
	      end;
	    end;
	    writeln(')');

{ cache }
	    if infotbl[$c]<>0 then writeln(' Cache line size ',infotbl[$c]*4,' Bytes (',infotbl[$c],' DWords)');

{ class code stuff }
	    write(' PCI Class ');

	    if infotbl[$b]=$ff then
	    begin
	      write('FFh ');
              textcolor(10);
	      write('(does not meet any PCI-SIG defined class)');
              textcolor(7);
	    end else
	    begin
	    
	    found:=false;
	    for i:=0 to high_class_name do
	    if infotbl[$b]=i then
	    begin
              textcolor(14);
	      write(PCI_class_names[i]);
              textcolor(7);
              found:=true;
	    end;
	    
	    if not found then
	    begin
	      textcolor(12);
	      write('Unknown! ');
	      textcolor(7);
	      write('(Class ',wrhex(infotbl[$b]),'h)')
	    end;
	    write(', type ');
	    found:=false;
	    for i:=0 to high_class_array do
	     begin
              if (pci_class_array[i].clase=infotbl[$b]) and
	      (pci_class_array[i].subclass=infotbl[$a]) and
	      (pci_class_array[i].progif=infotbl[$9]) then
	      begin
		found:=true;
                textcolor(14);
		write(PCI_class_array[i].name);
                textcolor(7);
	      end;
	    end;


	    if not found then
	    begin
	      for i:=0 to high_class_array do
	      begin
                if (pci_class_array[i].clase=infotbl[$b]) and
		(pci_class_array[i].subclass=infotbl[$a]) then
		begin
		  found:=true;
                  textcolor(14);
		  write(PCI_class_array[i].name);
                  textcolor(7);
		end;
	      end;
	    end;


	    if not found then
	    begin
              textcolor(12);
	      write('Unknown!');
	      textcolor(7);
	      write(' (Subclass ',wrhex(infotbl[$a]),'h, ProgIF ',wrhex(infotbl[9]),'h)');
	    end;
	    end;
	    writeln;
	  end;






	  if not summary then
	  begin
{ look for generic PCI IDE controller & decode it's info, if present }
	   if (infotbl[$b]=01) and (infotbl[$a]=01) then
	   begin
	     writeln(' PCI EIDE Controller Features :');
	     write('   BusMaster EIDE is ');
	     if infotbl[$9] and $80=0 then
	     begin
               textcolor(12);
	       write('NOT ');
               textcolor(7);
	     end;
	     writeln('supported');

	     write('   Primary   Channel is ');
	     if infotbl[$9] and 1=0 then
	     begin
	       writeln('at I/O Port 01F0h and IRQ 14');
	       if infotbl[$3c]<>14 then inc(irqmap[14]);
	     end else writeln('in native mode at Addresses 0 & 1');
	     write('   Secondary Channel is ');
	     if infotbl[$9] and 4=0 then
	     begin
	       writeln('at I/O Port 0170h and IRQ 15');
	       if infotbl[$3c]<>15 then inc(irqmap[15]);
	     end else writeln('in native mode at Addresses 2 & 3');
	   end;

	   end else
	   begin
{ summary mode: pick up IRQs only }
	     if (infotbl[$b]=01) and (infotbl[$a]=01) then
	     begin
	       if (infotbl[$9] and 1=0) and (infotbl[$3c]<>14) then inc(irqmap[14]);
	       if (infotbl[$9] and 4=0) and (infotbl[$3c]<>15) then inc(irqmap[15]);
	     end;
	   end;








{ if type 0 or 2 table & if Subsystem ID exists, display and scan file for match }

{ null out any old data }  
	   ss2c:=0;
	   ss2d:=0;
	   ss2e:=0;
	   ss2f:=0;
	   
	   if infotbl[$e] and $7f=0 then
	   begin
	     ss2c:=infotbl[$2c];
	     ss2d:=infotbl[$2d];
	     ss2e:=infotbl[$2e];
	     ss2f:=infotbl[$2f];
	   end;
	   
	   if infotbl[$e] and $7f=2 then
	   begin
	     ss2c:=infotbl[$40];
	     ss2d:=infotbl[$41];
	     ss2e:=infotbl[$42];
	     ss2f:=infotbl[$43];
	   end;
	   
	   
	   
	   
	   if (ss2c<>0) or (ss2d<>0) or (ss2e<>0) or (ss2f<>0) then
	   begin


{ subsystem ID }
	     write(' Subsystem ID ',wrhexw(ss2f shl 8+ss2e));
	     write(wrhexw(ss2d shl 8+ss2c),'h');
	     cmpstr:=wrhexw(ss2f shl 8+ss2e)+wrhexw(ss2d shl 8+ss2c);


	     genssid:=false;
	     if (ss2c=infotbl[0])
	     and (ss2d=infotbl[1])
	     and (ss2e=infotbl[2])
	     and (ss2f=infotbl[3]) then genssid:=true;

	     oemidnum:='';
	     oemidstr:='';
	     bogusid:=false;
	     failed:=true;
{ use the line that was read for revchk, or the first O or X entry will be missed! }
	     userev:=true;


	     if not eof(f) then
	     begin
	       repeat
		 if userev and (revchk<>'') then vstr:=revchk else readln(f,vstr);
		 userev:=false;

{ OEM Vendor ID }
		 if vstr[1]='O' then
		 begin
		   if copy(vstr,3,4)=copy(cmpstr,5,4) then
		   begin
		     oemidstr:=copy(vstr,8,length(vstr)); { closest match }
		     oemidnum:=copy(vstr,3,4); { matching vendor name }
		   end;
		 end;


		 if vstr[1]='S' then
		 begin
		   if copy(vstr,3,4)=copy(cmpstr,1,4) then
		   begin
		     if oemidnum<>'' then
		     begin
		       oemidstr:=copy(vstr,8,length(vstr));
		       begin
                         textcolor(14);
			 write(' ',oemidstr);
			 if genssid then
			 begin
                           textcolor(11);
			   writeln(' (Generic ID)')
			 end else writeln;
			 failed:=false;
                         textcolor(7);
		       end;
		     end;
		   end;
		 end;





{ Oddball 8 digit entry }
		 if (vstr[1]='X') and (copy(vstr,3,8)=cmpstr) then
		 begin
		   oemidnum:=copy(vstr,7,4); { matching vendor name }
		   bogusid:=true;
                   textcolor(14);
		   write(' ',copy(vstr,12,length(vstr)));
		   if genssid then
		   begin
                     textcolor(11);
		     writeln(' (Generic ID)')
		   end else writeln;
		   failed:=false;
                   textcolor(7);
		 end;


{ remember to ignore comment lines here also!!! }

	       until eof(f) or not failed or ((vstr[1]<>'O') and (vstr[1]<>'X') and (vstr[1]<>'S') and(vstr[1]<>';'));
	     end;

	     if failed then
	     begin
	       if oemidstr<>'' then
	       begin
                 textcolor(14);
		 write(' ',oemidstr);
                 textcolor(15);
		 write(' (Guess Only!)');
                 textcolor(7);
	       end else
	       begin
                 textcolor(12);
		 write(' Unknown');
	       end;

	       if genssid then
	       begin
                 textcolor(11);
		 writeln(' (Generic ID)')
	       end else writeln;
               textcolor(7);
	     end;


{ subsystem vendor }
	     write(' Subsystem Vendor ',wrhexw(ss2d shl 8+ss2c),'h');

	     if bogusid then
	     begin
               textcolor(15);
	       writeln(' Known Bad Subsystem ID - no Vendor ID Available');
               textcolor(7);
	     end else
	     begin
	       if oemidnum<>'' then cmpstr:=oemidnum
	       else cmpstr:=wrhexw(ss2d shl 8+ss2c);
	       close(f);		{ get back to start of file, as the}
	       reset(f);		{ subsys vendor may be higher up...!}
	       failed:=true;
	       if not eof(f) then
	       begin
		 repeat
		   readln(f,vstr);
                   if length(vstr)<3 then vstr:=';   oops';
		   if (vstr[1]='V') and (copy(vstr,3,4)=cmpstr) then
		   begin
                     textcolor(14);
		     writeln(' ',copy(vstr,8,length(vstr)));
		     failed:=false;
                     textcolor(7);
		   end;
		 until eof(f) or not failed;
	       end;
	       if failed then
	       begin
                 textcolor(12);
		 writeln(' Unknown');
                 textcolor(7);
	       end;
	     end;
	   end;
{ always }
	   close(f);











{ Memory & I/O  registers }
	   if not summary then
	   begin
{ type 0 header = 6 entries, type 1 = 2, type 2 = skip }
	   pp:=0;
	   if infotbl[$e] and $7f=0 then pp:=5;
	   if infotbl[$e] and $7f=1 then pp:=1;

	   if pp>0 then for nn:=0 to pp do
	   begin
	     if infotbl[$10+(nn*4)]+infotbl[$11+(nn*4)]+
	       infotbl[$12+(nn*4)]+infotbl[$13+(nn*4)]<>0 then
	     begin
	       write(' Address ',nn,' is a');
	       if infotbl[$10+(nn*4)] and 1=1 then
	       begin
		 write('n I/O Port : ');
		 addr:=infotbl[$13+(nn*4)] shl 8 + infotbl[$12+(nn*4)];
		 write(wrhexw(addr));
		 addr:=infotbl[$11+(nn*4)] shl 8 + (infotbl[$10+(nn*4)] and $fc);
		 write(wrhexw(addr),'h');
	       end else
	       begin
		 write(' Memory Address');
		 if infotbl[$10+(nn*4)] and 6=0 then write(' (anywhere in 0-4Gb');
		 if infotbl[$10+(nn*4)] and 6=2 then write(' (below 1Mb');
		 if infotbl[$10+(nn*4)] and 6=4 then write(' (anywhere in 64-bit space');
		 if infotbl[$10+(nn*4)] and 6=6 then write(' (reserved');
		 if infotbl[$10+(nn*4)] and 8=8 then write(', Prefetchable) : ') else write(') : ');
		 addr:=infotbl[$13+(nn*4)] shl 8 + infotbl[$12+(nn*4)];
		 write(wrhexw(addr));
		 addr:=infotbl[$11+(nn*4)] shl 8 + (infotbl[$10+(nn*4)] and $f0);
		 write(wrhexw(addr)+'h');
	       end;

{ size the register ?? }

	       writeln;
	     end;
	   end;
	   end;





{ all header types - list IRQ, if present }
           if (infotbl[$3c]>0) and (infotbl[$3c]<255) then
	   begin
	     write(' System IRQ ',infotbl[$3c],', INT# ');
	     if infotbl[$3d]=0 then write('-') else write(chr(infotbl[$3d]+64));
	     writeln;
	     inc(irqmap[infotbl[$3c]]);
             if infotbl[$3c]>15 then apic:=true;
	   end;




{ type 0,1 header only - List ExpROM, if present }
	   if not summary then
	   begin
	     if (infotbl[$e] and $7f=0) or (infotbl[$e] and $7f=1) then
	     begin
	       if infotbl[$e] and $7f=0 then lb:=$30;
	       if infotbl[$e] and $7f=1 then lb:=$38;

{ backup original value }
               rom_backup:=infotbl[lb+3] shl 24 + infotbl[lb+2] shl 16 + infotbl[lb+1] shl 8 + infotbl[lb];
{ perform test write }
               write_dword_hw(deviceid,func,bus,lb,$ffff,$fffe);
	       for i:=lb to lb+3 do infotbl[i]:=lookup_hw(deviceid,func,bus,i);
{ restore original value regardless }
               write_dword_hw(deviceid,func,bus,lb,rom_backup shr 16,rom_backup and $ffff);
{ if any bit in 31..11 = 1, we have expROM }
	       if (infotbl[lb+3]<>0) or (infotbl[lb+2]<>0) or ((infotbl[lb+1] and $f8)<>0) then
	       begin
{ find lowest 1-bit, gives expROM size }
	         romsize:=1;
	         romresult:=infotbl[lb+3] shl 24 + infotbl[lb+2] shl 16 + ((infotbl[lb+1] and $f8) shl 8);
	         found:=false;
	         repeat
	           if romresult and romsize=romsize then found:=true else romsize:=romsize shl 1;
	         until found or (romsize=0);
	         romsize:=romsize shr 10;
	         write(' Expansion ROM of ');
	         if romsize>1023 then write(romsize shr 10,'Mb') else write(romsize,'Kb');
	         write(' decoded by this card (Currently ');
	         if rom_backup and 1=1 then writeln('enabled)') else writeln('disabled)');
	       end;
{ restore infotbl[] to original values }
	       infotbl[lb]:=rom_backup and $ff;
	       infotbl[lb+1]:=(rom_backup shr 8) and $ff;
	       infotbl[lb+2]:=(rom_backup shr 16) and $ff;
	       infotbl[lb+3]:=(rom_backup shr 24) and $ff;
	     end;
	   end;









{ PCI Bridges info starts here }

{ type 1 header only - List bridge info }
	   if not summary then
	   begin

	   if infotbl[$e] and $7f=1 then
	   begin
	     writeln(' PCI Bridge Information:');
	     write('   Primary Bus Number ');
	     textcolor(11);
	     write(infotbl[$18]);
	     textcolor(7);
	     write(', Secondary Bus Number ');
	     textcolor(11);
	     write(infotbl[$19]);
	     textcolor(7);
	     write(', Subordinate Bus Number ');
	     textcolor(11);
	     writeln(infotbl[$1a]);
	     textcolor(7);

{ seconday bus command }
	     first:=true;
	     write('   Secondary Bus Command ',wrhexw(infotbl[$3f] shl 8 + infotbl[$3e]),'h ');
	     write('(');
	     if infotbl[$3e] and 1=1 then printstatus('parity detection');
	     if infotbl[$3e] and 4=4 then printstatus('ISA mapping');
	     if infotbl[$3e] and 8=8 then printstatus('VGA mapping');
	     if infotbl[$3e] and 32=32 then printstatus('master abort mode');
	     if infotbl[$3e] and 64=64 then printstatus('secondary bus is in RESET');
	     if infotbl[$3e] and 128=128 then printstatus('back-to-back transactions');
	     writeln(')');


{ secondary bus status }
	    write('   Secondary Bus Status ',wrhexw(infotbl[$1f] shl 8+infotbl[$1e]),'h');
	    if (infotbl[$1e]<>0) or (infotbl[$1f]<>0) then
	    begin
	      first:=true;
	      write(' (');
	      if infotbl[$1e] and 32=32 then printstatus('Supports 66MHz');
	      if infotbl[$1e] and 64=64 then printstatus('Supports UDF');
	      if infotbl[$1e] and 128=128 then printstatus('Supports Back-To-Back Trans.');
	      if infotbl[$1f] and 1=1 then printstatus('Data parity Error Detected');
	      if infotbl[$1f] and 8=8 then printstatus('Signalled Target Abort');
	      if infotbl[$1f] and 16=16 then printstatus('Received Target Abort');
	      if infotbl[$1f] and 32=32 then printstatus('Received Master Abort');
	      if infotbl[$1f] and 64=64 then printstatus('Received System Error');
	      if infotbl[$1f] and 128=128 then printstatus('Detected Parity Error');
	      case ((infotbl[$1f] and 6) shr 1) of
		0 : printstatus('Fast Timing');
		1 : printstatus('Medium Timing');
		2 : printstatus('Slow Timing');
		3 : printstatus('Unknown Timing');
	      end;
	      write(')');
	    end;
	    writeln;

{ latency }
	    writeln('   Secondary Bus Latency ',wrhex(infotbl[$1b]),'h');

{ I/O port range passed by bridge }
	     if (infotbl[$1c]<>0) or (infotbl[$1d]<>0) then
	     begin
	       write('   I/O Port Range Passed to Secondary Bus : ');

{ 16-bit I/O }
	       if infotbl[$1c] and $f=0 then
	       begin
		 if infotbl[$1d]<infotbl[$1c] then writeln('None') else
		 begin
		   write(wrhexb(infotbl[$1c] shr 4),'000h to ');
		   writeln(wrhexb(infotbl[$1d] shr 4),'FFFh (16-bit I/O space)');
		 end;
	       end;
{ 32-bit I/O }
	       if infotbl[$1c] and $f=1 then
	       begin
		 if (infotbl[$33] shl 8 + infotbl[$32]) < (infotbl[$31] shl 8 + infotbl[$30]) then writeln('None') else
		 begin
		   write(wrhexw(infotbl[$31] shl 8 + infotbl[$30]),wrhexb(infotbl[$1c] shr 4),'000h to ');
		   writeln(wrhexw(infotbl[$33] shl 8 + infotbl[$32]),wrhexb(infotbl[$1d] shr 4),'FFFh (32-bit I/O space)');
		 end;
	       end;
	     end;



{ memory range passed by bridge }
	     write('   Memory   Range Passed to Secondary Bus : ');
	     mb:=((infotbl[$21] shl 8) + infotbl[$20]) and $fff0;
	     ml:=((infotbl[$23] shl 8) + infotbl[$22]) or $000f;
	     if ml<mb then writeln('None') else writeln(wrhexw(mb),'0000h to ',wrhexw(ml),'FFFFh');


{ optional: prefetchable memory range passed by bridge }
	     if ((infotbl[$27] shl 8 + infotbl[$26])<>0) or ((infotbl[$25] shl 8 + infotbl[$24])<>0) then
	     begin
	       write('   Prefetchable Memory Range Passed to Secondary Bus : ');
	       mb:=((infotbl[$25] shl 8) + infotbl[$24]) and $fff0;
	       ml:=((infotbl[$27] shl 8) + infotbl[$26]) or $000f;
               if ml<mb then writeln('None') else
               begin
{ 64 bit base reg? }               
                 if infotbl[$24] and $0f=1 then write(wrhexw(infotbl[$2b] shl 8 + infotbl[$2a]),wrhexw(infotbl[$29] shl 8 + infotbl[$28]));
	         write(wrhexw(mb),'0000h to ');
{ 64 bits limit reg? }	       
                 if infotbl[$26] and $0f=1 then write(wrhexw(infotbl[$2f] shl 8 + infotbl[$2e]),wrhexw(infotbl[$2d] shl 8 + infotbl[$2c]));
	         writeln(wrhexw(ml)+'FFFFh');
	       end;
	     end;
	   end;
	   end;





{ type 2 header only - List bridge info }
	   if not summary then
	   begin
	     if infotbl[$e] and $7f=2 then
	     begin
	       writeln(' CardBus Bridge Information:');
	       write('   PCI Bus Number ');
	       textcolor(11);
	       write(infotbl[$18]);
	       textcolor(7);
	       write(', CardBus Bus Number ');
	       textcolor(11);
	       write(infotbl[$19]);
  	       textcolor(7);
	       write(', Subordinate Bus Number ');
	       textcolor(11);
	       writeln(infotbl[$1a]);
	       textcolor(7);

   	       write('   CardBus Bus Status ',wrhexw(infotbl[$17] shl 8+infotbl[$16]),'h');
    	       if (infotbl[$16]<>0) or (infotbl[$17]<>0) then
    	       begin
    	         first:=true;
    	         write(' (');
    	         if infotbl[$16] and 32=32 then printstatus('Supports 66MHz');
    	         if infotbl[$16] and 64=64 then printstatus('Supports UDF');
    	         if infotbl[$16] and 128=128 then printstatus('Supports Back-To-Back Trans.');
    	         if infotbl[$17] and 1=1 then printstatus('Data parity Error Detected');
    	         if infotbl[$17] and 8=8 then printstatus('Signalled Target Abort');
    	         if infotbl[$17] and 16=16 then printstatus('Received Target Abort');
    	         if infotbl[$17] and 32=32 then printstatus('Received Master Abort');
    	         if infotbl[$17] and 64=64 then printstatus('Received System Error');
    	         if infotbl[$17] and 128=128 then printstatus('Detected Parity Error');
    	         case ((infotbl[$17] and 6) shr 1) of
    		   0 : printstatus('Fast Timing');
    		   1 : printstatus('Medium Timing');
    		   2 : printstatus('Slow Timing');
    		   3 : printstatus('Unknown Timing');
    	         end;
    	         write(')');
    	       end;
    	       writeln;

    	       writeln('   CardBus Latency ',wrhex(infotbl[$1b]),'h');


{ ExCA BAR, should always be memory mapped, but handle I/O mapped just in case }
	     write('   ExCA Base ');
	     if infotbl[$10] and 1=0 then
	     begin
	       write('Memory Address : ');
  	       cbbase:=((infotbl[$13] shl 24) + (infotbl[$12] shl 16) + (infotbl[$11] shl 8)) and $fffff000;
	       writeln(wrhexl(cbbase),'h');
	     end else
	     begin
                write('I/O Address : Unknown (');
                cbbase:=((infotbl[$13] shl 24) + (infotbl[$12] shl 16) + (infotbl[$11] shl 8) + infotbl[$10]);
                writeln(wrhexl(cbbase),'h)!!');
             end;

    	       
{ memory ranges passed by bridge }
	     write('   Memory   Range 0 Passed to Secondary Bus : ');
	     cbbase:=((infotbl[$1f] shl 24) + (infotbl[$1e] shl 16) + (infotbl[$1d] shl 8)) and $fffff000;
	     cblimit:=((infotbl[$23] shl 24) + (infotbl[$22] shl 16) + (infotbl[$21] shl 8)) or $fff;
	     if cblimit<cbbase then writeln('None') else writeln(wrhexl(cbbase),'h to ',wrhexl(cblimit),'h');

	     write('   Memory   Range 1 Passed to Secondary Bus : ');
	     cbbase:=((infotbl[$27] shl 24) + (infotbl[$26] shl 16) + (infotbl[$25] shl 8)) and $fffff000;
	     cblimit:=((infotbl[$2b] shl 24) + (infotbl[$2a] shl 16) + (infotbl[$29] shl 8)) or $fff;
	     if cblimit<cbbase then writeln('None') else writeln(wrhexl(cbbase),'h to ',wrhexl(cblimit),'h');

{ i/o ranges passed by bridge }
	     write('   I/O Port Range 0 Passed to Secondary Bus : ');
	     cbbase:=((infotbl[$2f] shl 24) + (infotbl[$2e] shl 16) + (infotbl[$2d] shl 8) + (infotbl[$2c] and $fc));
	     cblimit:=((infotbl[$33] shl 24) + (infotbl[$32] shl 16) + (infotbl[$31] shl 8) + (infotbl[$30] or 3));
	     if infotbl[$2c] and 3=1 then writeln(wrhexl(cbbase),'h to ',wrhexl(cblimit),'h (32-bit)') else writeln(wrhexw(cbbase and $ffff),'h to ',wrhexw(cblimit and $ffff),'h (16-bit)');

	     write('   I/O Port Range 1 Passed to Secondary Bus : ');
	     cbbase:=((infotbl[$37] shl 24) + (infotbl[$36] shl 16) + (infotbl[$35] shl 8) + (infotbl[$34] and $fc));
	     cblimit:=((infotbl[$3b] shl 24) + (infotbl[$3a] shl 16) + (infotbl[$39] shl 8) + (infotbl[$38] or 3));
	     if infotbl[$34] and 3=1 then writeln(wrhexl(cbbase),'h to ',wrhexl(cblimit),'h (32-bit)') else writeln(wrhexw(cbbase and $ffff),'h to ',wrhexw(cblimit and $ffff),'h (16-bit)');

    	       
    	     end;
    	   end;



{ explore the capabilities list, if present }
	   if not summary then
	   begin
	     if (infotbl[6] and $10=$10) then docapdecode;
	   end;


{ do a hex-dump, if requested }
	   if dumpregs then
	   begin
	     writeln;
	     writeln(' Hex-Dump of device configuration space follows:');
	     write('  0000  ');
	     for i:=0 to $ff do
	     begin
	       if (i>0) and (i mod 16=0) then
	       begin
		 write('   ');
		 for j:=i-16 to i-1 do if ord(infotbl[j])<32 then write('.') else write(chr(infotbl[j]));
		 writeln;
		 write('  ',wrhexw(i),'  ');
	       end;
	       write(wrhex(infotbl[i]),' ');
	     end;
	     write('   ');
	     for j:=240 to 255 do if ord(infotbl[j])<32 then write('.') else write(chr(infotbl[j]));
	     writeln;
	   end;

    writeln;
  end;
end;


{ Start NT stuff }


procedure initNT_system;
var result : dword;

begin
  result:=gwiopm_driver.openscm;
  if result<>error_success then writeln('NT: Open SCM Error : ',gwiopm_driver.errorlookup(result));
  
  result:=gwiopm_driver.install('');
  if result<>error_success then
  begin
    writeln('NT: Driver Problem detected: attempting to fix...');
    gwiopm_driver.stop;
    gwiopm_driver.remove;
    writeln('NT: Driver hopefully fixed, retrying...');
    result:=gwiopm_driver.install('');
  end;

  if result<>error_success then
  begin
    writeln('NT: Couldn''t Install/Start driver, error : ',gwiopm_driver.errorlookup(result),' : PCI32 halted!');
    writeln('NT: Ensure you have administrative rights, that GWIOPM.SYS is present, and that PCI32 is run from a local drive');
    halt(12);
  end;

  gwiopm_driver.start;
  gwiopm_driver.closescm;
  gwiopm_driver.DeviceOpen;

  GWIOPM_Driver.LIOPM_Set_Ports($0cf8, $0cff, true);
  GWIOPM_Driver.IOCTL_IOPMD_ACTIVATE_KIOPM;
end;



procedure killNT_system;
begin
  GWIOPM_Driver.LIOPM_Set_Ports($cf8, $cff, false);
  GWIOPM_Driver.IOCTL_IOPMD_ACTIVATE_KIOPM;

  gwiopm_driver.DeviceClose;
  gwiopm_driver.openscm;
  gwiopm_driver.stop;
  gwiopm_driver.remove;
  gwiopm_driver.closescm;
end;






procedure workoutbusses;
var 
  i,
  k,
  bus,
  index,
  func,
  deviceid	: integer;
  gotone	: boolean;
  

begin
  count:=0;
  empty:=0;

{ 0=un-identified, 1=PCI, 2=AGP, 3=PCIe, 4=CardBus, 5=PCI-X  6=empty bus }
  for i:=0 to 255 do
  begin
    bustype[i]:=0;
    busvalid[i]:=false;
  end;

  for bus:=0 to 255 do
  begin
    gotone:=false;
    if debug then write('Examining bus ',bus);
    for deviceid:=0 to $1f do
    begin
      func:=0;
      repeat
        index:=0;
        found:=false;
        repeat
          info:=lookup_hw(deviceid,func,bus,index);
    	  infotbl[index]:=info;
          inc(index);
  	  if index=2 then if (infotbl[0]=$ff) and (infotbl[1]=$ff) then index:=$100 else found:=true;
        until (index=$100); { index }
        
      
      
        if found then
        begin
{ we have a device ! }
          if debug then if not gotone then writeln;
          gotone:=true;
{ mark this as a valid bus }
          busvalid[bus]:=true;
{ update highest known bus, if relevant }
          if bus>pci_hibus then pci_hibus:=bus;
{ update the tree of devices for the showtree function while we're here }
          tree[count].b:=bus;
          tree[count].d:=deviceid;
          tree[count].f:=func;
          tree[count].vid:=infotbl[1] shl 8 + infotbl[0];
          tree[count].did:=infotbl[3] shl 8 + infotbl[2];
          tree[count].cl:=infotbl[$b];
          tree[count].su:=infotbl[$a];
          tree[count].pi:=infotbl[$9];


{ look for a bridge : type 1 or 2, remember it }
          if (infotbl[$e] and $7f=1) or (infotbl[$e] and $7f=2) then
          begin
            tree[count].secbus:=infotbl[$19];
            tree[count].subbus:=infotbl[$1a];
{ !! remember which busses are valid, if found here too. Bus may be empty, but is still valid since the bridge says so }            
            for k:=tree[count].secbus to tree[count].subbus do
            begin
              busvalid[k]:=true;
              if debug then
              begin
                write('   Sub-bus ');
                textcolor(11);
                write(k);
                textcolor(7);
                writeln(' marked valid');
              end;
            end;
{ !! update highest bus number }
            if infotbl[$1a]>pci_hibus then pci_hibus:=infotbl[$1a];
          end else tree[count].secbus:=0;
          inc(count);
          if debug then
          begin
            write(' Device Found : ',count-1,' B:',tree[count-1].b,' D:',tree[count-1].d,' F:',tree[count-1].f);
            writeln(' VID:',wrhexw(tree[count-1].vid),' DID:',wrhexw(tree[count-1].did));
          end;
{ end tree update }

{ look for cardbus by finding bridge type 2 }
          if (infotbl[$e] and $7f=2) then
          begin
           if debug then writeln('Cardbus Sniffed: B:',bus,' D:',deviceid,' F:',func,' supposed bus:',infotbl[$19]);
           if bustype[infotbl[$19]]=0 then bustype[infotbl[$19]]:=4 else 
           if bustype[infotbl[$19]]=2 then
           begin
             if debug then writeln('Over-riding AGP : ',busnames[bustype[infotbl[$19]]],' ID''d as CardBus');
             bustype[infotbl[$19]]:=4;
           end else
           if debug then writeln('  Conflicting Bus ID : Bus ',infotbl[$19],' is ',busnames[bustype[infotbl[$19]]],' ID''d as CardBus (Ignored)');
          end;


{ look for pci-x bridge by finding bridge type header and pci-x msi on that device }
          if (infotbl[$e] and $7f=1) and (infotbl[$34]<>0) then
          begin
            cap_ptr:=infotbl[$34];
            repeat
              if infotbl[cap_ptr]=07 then
              begin
                if debug then writeln('PCI-X Sniffed: B:',bus,' D:',deviceid,' F:',func,' supposed bus:',infotbl[$19]);
                if bustype[infotbl[$19]]=0 then bustype[infotbl[$19]]:=5 else
                if bustype[infotbl[$19]]=2 then
                begin
                  if debug then writeln('Over-riding AGP : ',busnames[bustype[infotbl[$19]]],' ID''d as PCI-X');
                  bustype[infotbl[$19]]:=5;
                end else
                if debug then writeln('  Conflicting Bus ID : Bus ',infotbl[$19],' is ',busnames[bustype[infotbl[$19]]],' ID''d as PCI-X (Ignored)');
              end;
              cap_ptr:=infotbl[cap_ptr+1];
            until cap_ptr=0;
          end;
            



	
{ look for PCI Express by existance of PCIe capability}
{type 0}  if infotbl[$e] and $7f=0 then cap_ptr:=infotbl[$34];
{type 1}  if infotbl[$e] and $7f=1 then cap_ptr:=infotbl[$34];
{type 2}  if infotbl[$e] and $7f=2 then cap_ptr:=infotbl[$14];
          if cap_ptr<>0 then
          repeat
            if infotbl[cap_ptr]=$10 then
            begin
              if debug then writeln('PCIe Sniffed: B:',bus,' D:',deviceid,' F:',func,' supposed bus:',bus);
              if bustype[bus]=0 then bustype[bus]:=3 else
              if bustype[bus]=2 then
              begin
                if debug then writeln('Over-riding AGP : ',busnames[bustype[bus]],' ID''d as PCI Express');
                bustype[bus]:=3;
              end else
              if debug then writeln('  Conflicting Bus ID : Bus ',bus,' is ',busnames[bustype[bus]],' ID''d as PCI Express (Ignored)');
            end;
            cap_ptr:=infotbl[cap_ptr+1];
          until cap_ptr=0;
  
{ try to guess AGP bus number by looking for a set VGA Mapping flag on a PCI-PCI Bridge }
	  if infotbl[$e] and $7f=1 then if infotbl[$3e] and 8=8 then
	  begin
            if debug then writeln('AGP Sniffed: B:',bus,' D:',deviceid,' F:',func,' supposed bus:',infotbl[$19]);
            if bustype[infotbl[$19]]=0 then bustype[infotbl[$19]]:=2 else 
            if debug then writeln('  Conflicting Bus ID : Bus ',infotbl[$19],' is ',busnames[bustype[infotbl[$19]]],' ID''d as AGP (Ignored)');
	  end;

        end;  { of found }

      inc(func);
{ if func 0 = invalid device, don't test for presence of func 1->7 at all. [$e] isn't valid if [0] and [1] aren't!! }
      if (func=1) and (infotbl[0]=$ff) and (infotbl[1]=$ff) then func:=8;
{ If not multi-device device, then don't test for func 1-7 as some cards
incorrectly answer back on all 8 function numbers!!! S3 trio64, for example - stupid!  }    
      if (func=1) and (infotbl[$e] and $80=0) then func:=8;
      until func=8; {func }
    end; {dev }





{ flag empty busses, but not if already positively known as something else! }
{ nb: only CardBus can be positively ID'd even when empty }
    if not gotone and (bustype[bus]<>4) and busvalid[bus] then
    begin
      bustype[bus]:=6;
      if debug then writeln(' -> Bus ',bus,' flagged as empty');
      tree[count].vid:=$ffff;
      tree[count].did:=$ffff;
      tree[count].b:=bus;
      inc(count);
    end else if debug and not gotone then writeln;
 


  end; { bus }
  
  
  

{ finally, if it's a bus we've seen, AND not already ID'd, ID it as plain old PCI... }
  for i:=0 to 255 do if busvalid[i] and (bustype[i]=0) then bustype[i]:=1;

 
 
  if debug then
  begin
    writeln;
    writeln('pci_hibus : ',pci_hibus);
    writeln('busses ID''d as follows:');
    for i:=0 to 255 do
    begin
      if busvalid[i] then writeln(' Bus ',i,' : valid : ',busnames[bustype[i]]);
    end;
    writeln('Count : ',count);
    writeln;
  end;
end;




{ final summarial IRQ info }
procedure doirqsummary;
var
  i	: integer;
  
begin
  writeln;
  write('IRQ Summary: ');
  failed:=true;
  disp:=0;
  for i:=0 to 255 do if irqmap[i]>0 then inc(disp); { count IRQs }
  for i:=0 to 255 do if irqmap[i]>0 then
  begin
    if failed then
    begin
      if disp=1 then write('IRQ ') else write('IRQs ');
    end else write(',');
    write(i);
    failed:=false;
  end;
  if failed then writeln('No IRQ''s are used by PCI Devices!') else
  begin
    if disp=1 then write(' is') else write(' are');
    writeln(' used by PCI devices');
  end;

  write('Shared IRQs: ');
  failed:=true;
  for i:=0 to 255 do if irqmap[i]>1 then
  begin
    if not failed then write('             ');
    write('IRQ ',i,' is shared by ',irqmap[i]);
    writeln(' PCI Devices');
    failed:=false;
  end;
  if failed then writeln('There are no shared PCI IRQs');
  if apic then
  begin
    writeln;
    writeln('IRQ control is currently managed by the system APIC controller - IRQ info is');
    writeln('not actual hardware settings...');
  end;
  writeln;
end;



{ draw the bus/device/function tree }
procedure dobdftree;
var
  i,
  j	: integer;
  
begin
  writeln;
  writeln('PCI Busses, Devices and Device Functions Tree');
  olb:=$ff;
  
  if ioredirected then
  begin
    li[1]:='--';
    li[2]:='|--';
    li[3]:='+--';
    li[4]:='| ';
  end else
  begin
    li[1]:='ÄÄ';
    li[2]:='ÀÄÄ';
    li[3]:='ÃÄÄ';
    li[4]:='³ ';
  end;  
    
  for i:=count to 511 do
  begin
    tree[i].d:=$ff;											{ flag unused entries }
    tree[i].b:=$ff;
  end;
   
  for i:=0 to count-1 do
  begin
    if tree[i].b<>olb then
    begin
      writeln;
      write(li[1],'Bus ');
      textcolor(11);
      write(tree[i].b);
      textcolor(7);
      if bustype[tree[i].b]=6 then writeln else writeln(' (',busnames[bustype[tree[i].b]],')');								{ bus }
    end;

    
    notlast:=false;												{ last device this bus ?? }
    for j:=i+1 to count-1 do
    begin
      if tree[j].b=tree[i].b then if tree[j].d<>tree[i].d then notlast:=true;
    end;
    
    if not notlast and (tree[i].f=0) then write('   ',li[2]) else						{ final tree DEV item for each bus}
    if tree[i].f=0 then write('   ',li[3]) else if notlast then write('   ',li[4],' ') else write('      ');	{ single F or MF dev }
 
{ up to here colums before "device" }
{ after here colums for func }

    if (tree[i].vid=$ffff) and (tree[i].did=$ffff) then
    begin
      writeln(' No Devices Currently Present');
      inc(empty);
    end else
    begin

    if tree[i].f=0 then write(' Device ',wrhex(tree[i].d),'h      ');					{ device }
    if (tree[i].f=0) and (tree[i+1].f>0) then
    begin
      writeln('');		  									{ MF func 0 only }
      write('   ');
      if tree[i+1].b<>tree[i].b then write('  ') else if notlast then write(li[4]) else write('  ');	{ draw D bar is not last D on this bus }
      write('    ',li[3],' Function ',tree[i].f);
    end;
    if tree[i].f>0 then if tree[i+1].f<tree[i].f then write('   ',li[2],' Function ',tree[i].f) else write('   ',li[3],' Function ',tree[i].f);		{ func bar }
    write('  ',wrhexw(tree[i].vid),'h:',wrhexw(tree[i].did),'h  ');					{ VID:DID in hex }


{ display device class info }
    found:=false;
    for j:=0 to high_class_array do
    begin
      if (pci_class_array[j].clase=tree[i].cl) and
      (pci_class_array[j].subclass=tree[i].su) and
      (pci_class_array[j].progif=tree[i].pi) then
      begin
        write(PCI_class_array[j].name);
        found:=true;
      end;
    end;
    if not found then
    for j:=0 to high_class_array do
    begin
      if (pci_class_array[j].clase=tree[i].cl) and
      (pci_class_array[j].subclass=tree[i].su) then
      begin
        found:=true;
        write(PCI_class_array[j].name);
      end;
    end;
    if not found then write('Unknown');
    
    for j:=0 to high_class_name do
    if tree[i].cl=j then write(' ',PCI_class_names[j]);
    
{ do pointers to next bus }    
    if tree[i].secbus>0 then
    begin
      write(' > Bus ');
      textcolor(11);
      write(tree[i].secbus);
      textcolor(7);
      if tree[i].subbus<>tree[i].secbus then
      begin
        write(' (..To Bus ');
        textcolor(11);
        write(tree[i].subbus);
        textcolor(7);
        write(')');
      end;
    end; { of 'no devices currently present' }
    writeln;
    end;
    
    olb:=tree[i].b;
  end;
  writeln;
  
  
  
{ mention all subsequent known CardBus busess with no devices. This should only happen for CardBus-es as other bus types }  
{ are picked up in the tree[] }

  for i:=olb+1 to pci_hibus do
  if (bustype[i]<>0) and (busvalid[i]=true) then
  begin
    write(li[1],'Bus ');
    textcolor(11);
    write(i);
    textcolor(7);
    writeln(' (',busnames[bustype[i]],')');
    writeln('   ',li[2],' No Devices Currently Present');
    writeln;
  end;
  writeln('A total of ',count-empty,' Devices Found'); { count is 0 based, so don't decrement here }
  writeln;
end;  







{ find out the OS type, just for interest's sake - program works the same regardless }

procedure getwinver;

var st1, st2 : string;

begin
  opsys:='';
  case Win32MajorVersion of
    4 : begin
          case Win32MinorVersion of
            0 : opsys:='Win NT 4.0';
          end;
        end;
    5: begin
         case Win32MinorVersion of
           0 : opsys:='Win 2000';
           1 : opsys:='Win XP';
           2 : opsys:='Win 2003';
         end;
       end;
    6: begin
         case win32minorversion of
           0 : opsys:='Win Vista';
         end;
       end;
  end;
  
  if opsys='' then
  begin
    str(Win32MajorVersion,st1);
    str(Win32MinorVersion,st2);
    opsys:='?? [ID:'+st1+'.'+st2+']';
  end;
  
end;  















begin
{ initialise important variables }
  debug:=false;
  businfo:=true;
  dumpregs:=false;
  summary:=false;
  installermode:=false;
  showtree:=false;
  for i:=0 to 255 do irqmap[i]:=0;
  failed:=true;
  textattr:=07;
  pci_hibus:=0;
  getwinver;

{ process commandline }  
  if paramcount>0 then
  begin
    for i:=1 to paramcount do
    begin
      cmdstr:=paramstr(i);
      for j:=1 to length(cmdstr) do cmdstr[j]:=upcase(cmdstr[j]);
      if (cmdstr='/Z') or (cmdstr='-Z') then debug:=true;
      if (cmdstr='/R') or (cmdstr='-R') then showtree:=true;
      if (cmdstr='/D') or (cmdstr='-D') then dumpregs:=true;
{      if (cmdstr='/B') or (cmdstr='-B') then businfo:=true;}	{ businfo now on by default }
      if (cmdstr='/S') or (cmdstr='-S') then summary:=true;
      if (cmdstr='/I') or (cmdstr='-I') then installermode:=true;
      if (cmdstr='?') or (cmdstr='/?') or (cmdstr='-?') then
      begin
	writeln;
	writeln(' Help for PCI32  (Version '+revision+')');
        textcolor(8);
	writeln('ÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ');
        textcolor(7);
	writeln;
        writeln('Usage: PCI32 [-D] [-S] [-B] [-I] [-R] [-?]  [] indicates optional parameter');
	writeln;
	writeln;
	writeln('-D : Do a hex-dump of each device''s entire configuration space');
	writeln('-S : Create a brief, summary report only; only devices and IRQs listed');
{	writeln('-B : Enable display of the Bus, Device & Function information');}
	writeln('-I : Installer mode: produce raw data dump (for use with auto-setup programs)');
	writeln('-R : Draw a Tree of Busses, Devices and Device Functions');
	writeln('-? : Displays this help screen!');
	writeln;
	writeln('PCI32 Supports generating reports to a file or printer using MS-DOS pipes; i.e.');
	writeln;
	writeln('  PCI32 -D > REPORT.TXT  (Saves report to file "REPORT.TXT")');
	writeln;
	writeln('PCI32 is written by Craig Hart, and is released as freeware with no restictions');
	writeln('on use, copying, modification, or sourcecode use in other software.');
	write('Visit ');
        textcolor(11);
	write('http://members.datafast.net.au/dft0802 ');
        textcolor(7);
	writeln('for updates to the program and the PCI Database file PCIDEVS.TXT');
	halt(10);
      end;
    end;
  end;

{ fix up conflicting commandline switches }
  if installermode then
  begin
    dumpregs:=false;
    businfo:=false;
    summary:=false;
  end;
  if summary then dumpregs:=false;

{ open pcidevs.txt }
  if not installermode then
  begin
    assign(f,extractfilepath(paramstr(0))+'pcidevs.txt');
    {$i-}
    reset(f);
    if ioresult<>0 then
    begin
      writeln('PCI Halted:');
      writeln;
      writeln('Sorry, my PCIDEVS.TXT datafile is missing!!!');
      halt(11);
    end;
    close(f);
    {$i+}
  end;

{ Setup kernel mode device driver }
  initNT_system;
   
{ display header message }
  if not installermode then
  begin
    writeln(' Craig Hart''s PCI+AGP bus sniffer, Version '+revision+', freeware made in 1996-2005.');
    writeln;
    writeln('Searching for Devices using CFG Mechanism 1 [OS: ',opsys,' ',Win32CSDVersion,']');
    writeln;
    writeln;
  end;

{ figure out bus type for each bus .. also, determine pci_hibus }
  workoutbusses;

{ run thru all devices and report the statistics on each }
  for bus:=0 to pci_hibus do
  begin
    for deviceid:=0 to $1f do
    begin
      func:=0;
      repeat
        index:=0;
	repeat
          info:=lookup_hw(deviceid,func,bus,index);
	  infotbl[index]:=info;
          inc(index);
{don't try to read cfg-space of non-existant devices: hangs some chipsets!}
	  if index=2 then if (infotbl[0]=$ff) and (infotbl[1]=$ff) then index:=$100;
{don't read past $3f if in short-info modes [except for cardbus: read up to $43]; avoids crashing on intolerant hardware!}
	  if index=$40 then if infotbl[$e] and $7f<>2 then if installermode or summary then index:=$100;
	  if index=$44 then if installermode or summary then index:=$100;
	until (index=$100); { index }
	if (infotbl[0]<>$ff) or (infotbl[1]<>$ff) then showallinfo;
        inc(func);
{ if func 0 = invalid device, don't test for presence of func 1->7 at all. [$e] isn't valid if [0] and [1] aren't!! }
        if (func=1) and (infotbl[0]=$ff) and (infotbl[1]=$ff) then func:=8;
{ If not multi-device device, then don't test for func 1-7 as some very very old cards incorrectly answer back on all 8 function numbers!!! S3 trio64, for example - stupid!  }
        if (func=1) and (infotbl[$e] and $80=0) then func:=8;
      until func=8; {func}
    end; {dev}
  end; {bus}

{ ok, main per-device report is done... now do anything else the user asked for }

{ irq info }
  if not installermode then doirqsummary;

{ bus/dev/func tree }
  if showtree and not installermode then dobdftree;

{ all done, so shutdown and quit }    
  killNT_system;
end.
