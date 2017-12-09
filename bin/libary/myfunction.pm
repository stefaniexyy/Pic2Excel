package myfunction;
use strict;
use warnings;
use 5.010;
use POSIX qw(strftime);
use Time::Local qw(timelocal_nocheck timelocal timegm);
use Spreadsheet::ParseExcel qw(FmtUnicode);
use Spreadsheet::Read qw(row ReadData cellrow);
use vars qw(@ISA @EXPORT %EXPORT_TAGS);
use Exporter;
@ISA=qw(Exporter);
@EXPORT=qw(timezone cal_time date_to_data cal_month current_time);

my $sysdate=strftime("%Y-%m-%d",localtime());
my $sysyear=strftime("%Y-",localtime());
my $summer=substr($sysyear,0,4)."0401000000";
my $winter=substr($sysyear,0,4)."1031000000";
my $summer_time=timegm('00','00','00','01','03',substr($sysyear,0,4));#timelocal的月份是从0开始的
my $winter_time=timegm('00','00','00','31','09',substr($sysyear,0,4));
my @s_day;
my @w_day;
foreach (0..37){
	my $t_year;
	length($_)==1?($t_year='200'.$_):($t_year='20'.$_);
	my $summer_time=timegm('00','00','00','01','03',$t_year);#timelocal的月份是从0开始的
	my $winter_time=timegm('00','00','00','31','09',$t_year);
	my ($s_sec, $s_min, $s_hour, $s_mday, $s_mon, $s_year, $s_wday, $s_yday, $s_isdst) = gmtime($summer_time);
	my ($w_sec, $w_min, $w_hour, $w_mday, $w_mon, $w_year, $w_wday, $w_yday, $w_isdst) = gmtime($winter_time);
	$s_day[$_]=$t_year."040".(1+(7-$s_wday))."020000";
	$w_wday==7 ? ($w_day[$_]=$t_year."1031020000") : ($w_day[$_]=$t_year."10".(31-$w_wday)."020000");

}

sub data_format{#把window的换行符转换为空格
	my $data=shift;
	$data=~ s/\r\n/ /g;
	return $data;
}

sub read_excel{#从实例化中间表读取数据
	# my $instance 文件位置
	# $sheet_num第几个工作表
	# $outfile 输出文件位置
	# $start 从第几行开始输出
	my($function_name,$instance,$sheet_num,$outfile,$start)=@_;
	my $book=ReadData($instance);
	my $sheet=$book->[$sheet_num];#设定工作表
	open(OUTPUT,'>'.$outfile)||die "error can't create outfile\n";
	my $ss = ReadData ($instance, attr => 1)->[$sheet_num];#设定读取参数
	foreach my $rownum($start..$ss->{maxrow}){#设定最大行
		my @row=cellrow ($sheet,$rownum);
		if(defined($row[0])){
			foreach my $arr_num(0..(scalar(@row)-1)){
				if(defined ($row[$arr_num])){#判断变量是否初始化
					print  OUTPUT  data_format($row[$arr_num])."|";
				}
			}
			print OUTPUT "\n";
		}elsif(defined($row[1])){
			foreach my $arr_num1(1..(scalar(@row)-1)){
				if(defined ($row[$arr_num1])){#判断变量是否初始化
					print  OUTPUT  data_format($row[$arr_num1])."|";
				}
			}		
			print OUTPUT "\n";
		}
	}
	close(OUTPUT);
}


sub timezone{#判断当前是夏令时还是冬令时 如果是夏令时则输出5，冬令时则输出6
	my ($input)=@_;
	$input=~ s/[^0-9]//g;
	my $switch;
	if(($input>$w_day[substr($input,2,2)])||($input<=$s_day[substr($input,2,2)])){
		$switch=6;
	}else{
		$switch=5;
	}
	return $switch;
}


sub cal_time{#时间的加减
	my ($date,$value)=@_;#输入的时间，小时数
	$date=~ s/[^0-9]/:/g;
	length($date)<14?($date=$date." 00:00:00"):($date=$date);
	my @date=split(/:/,$date,-1);
	my $time=timegm($date[5],$date[4],$date[3],$date[2],($date[1]-1),$date[0])+$value*3600;
	my $result=strftime "%Y-%m-%d %H:%M:%S",gmtime($time); 
	return $result;
}

sub date_to_data{#将时间转化为一串数字 支持2015-01-01 11:11:11 或者2015/11/11 11:11:11格式
	my ($input_time)=@_;
	my $data;
	length($input_time)<12?($input_time=$input_time." 00:00:00"):($input_time=$input_time);
	$input_time=~ s/[^0-9]//g;
	return $input_time;
}

sub cal_month{
	my($date,$month_num)=@_;
	$date=~ s/[^0-9]/:/g;
	my @month=(31,28,31,30,31,30,31,31,30,30,30,31);

	my @date=split(/:/,$date,-1);
	if($month_num>=0){
		$month_num+$date[1]>12?(($date[1],$date[0])=(($date[1]+$month_num-12)%12,($date[1]+$month_num-($date[1]+$month_num)%12)/12+$date[0])):($date[1]=$date[1]+$month_num);		
	}else{
		$month_num+$date[1]<1?(($date[1],$date[0])=(((12-abs($month_num)%12+abs($month_num))+$month_num+$date[1])%12,($date[1]+$month_num-(12+$date[1]+$month_num)%12)/12+$date[0])):($date[1]=$date[1]+$month_num);
	}
	if($date[1]==0){
		$date[1]=12;
		$date[0]-=1;
	}
	$date[0]%4==0?($month[1]=29):($month[1]=28);
	$date[2]>$month[$date[1]-1]?($date[2]=$month[$date[1]-1]):($date[2]=$date[2]);#如果月末超过了当前月的最大值，取当前月的最大值

	my $time=timegm($date[5],$date[4],$date[3],$date[2],($date[1]-1),$date[0]);

	my $result=strftime "%Y-%m-%d %H:%M:%S",gmtime($time); 
	return $result;
}

sub  current_time{ #获取系统当前时间
	my ($sec,$min,$hour,$mday,$mon,$year,$wday,$yday,$isdst) = localtime;
	$year += 1900;#年是从1900开始的
	$mon += 1;#月份是从0开始的
	my $datetime = sprintf ("%d-%02d-%02d %02d:%02d:%02d", $year,$mon,$mday,$hour,$min,$sec);
	return $datetime;
}
1;
__END__

=head1 AUthor

	Willy Xi

=head1 Version

	V0.1_build2_20170721

=head1 Descriptiom

=head2 read_excel

	ooperl use myfunction->read_excel()
	need 4 parameter: 1. xls file path 2:sheet num 3:output file path output 4:start line 
	example:myfunction->read_excel('../input/example.xls',1,'../output/example.dat',1);

=head2 date_to_data

	transfer a date format data to number
	2017-01-01 00:00:00->20170101000000 or 2017-01-01->20170101000000
	use myfunction qw(date_to_data);

=head2 cal_time

	calculate hours
	use myfunction qw(cal_time);
	cal_time('2017-01-01 00:00:00',5);# add 5 hour

=head2 timezone

	calcuate   UTC time  to mexico local time
	use myfunction qw(timezone);
	timezone('2016-01-01 00:00:00');#output 6

=head2 cal_month

	calculate month 
	use myfunction qw(cal_month);
	cal_month('2016-01-01 00:00:00',2);#add 2 months

=head2  current_time
	
		get current system time 
		
=cut