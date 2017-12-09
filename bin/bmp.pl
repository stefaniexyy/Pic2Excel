#!/usr/bin/perl
use strict;
use warnings;
use DBI;
use DBD::Oracle;
use 5.010;
use lib './libary';
use JSON;
use Excel::Writer::XLSX;

#                              _.._        ,----------------------.
#                           ,'      `.    ( This is my test script )
#                          /  __) __` \    `-,--------------------'
#                         (  (`-`(-')  ) _.-'
#                         /)  \  = /  (
#                        /'    |--' .  \
#                       (  ,---|  `-.)__`
#                        )(  `-.,--'   _`-.
#                       '/,'          (  Uu",
#                        (_       ,    `/,-' )
#                        `.__,  : `-'/  /`--'
#                          |     `--'  |
#                          `   `-._   /
#                           \        (
#                           /\ .      \.
#                          / |` \     ,-\
#                         /  \| .)   /   \
#                        ( ,'|\    ,'     
#                        | \,`.`--"/      }
#                        `,'    \  |,'    /
#                       / "-._   `-/      |
#                       "-.   "-.,'|     ;
#                      /        _/["---'""]
#                     :        /  |"-     '
#                     '           |      /
#                                 `      |
sub big_hex2dec{#超大16进制转10进制
    my $in=shift;
    my $in_length=length($in);
    my $result=0;
    for(1..$in_length){
         $result+=hex(substr($in,$_-1,1))*16**($in_length-$_);
    }    
    return $result;
}
sub reverse_order{#按照指定长度反向输出字符串
    my ($character,$size)=@_;#字符串，字符串长度
    my $length=length($character);
    my $real_length=($length)%($size)==0?($length):($length+$size-$length%$size);
    my ($result,$count)=("",0);
    for(1..$real_length/$size){
        if($length>($count+1)*$size){
            $result=substr($character,$count*$size,$size).$result;
            $count++;
        }else{
            $result=substr($character,$count*$size,($length-($count+1)*$size)+$size).$result;
        }
    }
    return $result;
}
sub bmp_24_bit{#处理24为bmp
    say 'pixel deep is 24bit';
    my ($file,$offset_value,$width,$height)=@_;#文件名，偏移值，图像宽，图像高
    open(BMP,$file)||die "error can not open $file,$!\n";
    seek(BMP,$offset_value,0)||die "error ,seek fail,$!\n";#移动到像素开始
    open(DEBUG,'>./bmp.dat')||die "error can not  create debug file,$!";
    my ($actual_file_name)=($file)=~/(\w+)\./;
    my $real_width=$width%4 ne 0?(4-$width%4 +$width):($width);#bmp 是每行存储的字节数是4的倍数，如果不是4的倍数需要补0
    my $workbook=Excel::Writer::XLSX->new('./'.$actual_file_name.'.xlsx');
    my $worksheet=$workbook->add_worksheet('sheet1');
    $worksheet->set_column(0,$width,2);#把excel的单元格弄成正方形，因为像素是方的
    my $buf;
    ############
    for(1..$height){#纵向像素
        my $height_count=$_;
        #say 'current heigth:'.$height_count;
        print DEBUG "\n";
        for(1..$width){#横向像素
            my $width_count=$_;
            next unless $width_count<=$width;
            read(BMP,$buf,3);#每次读取3个字节，BGR
            my $color=unpack("H*",$buf);
            $color=reverse_order($color,2);
            my $format=$workbook->add_format();
            $format->set_bg_color('#'.$color);
            print DEBUG '|'.$color;
            $worksheet->write($height-$height_count,$width_count,' ',$format);
        }
        if($real_width-$width>2){
            read(BMP,$buf,(4-$real_width+$width));       
        }else{
            read(BMP,$buf,($real_width-$width));
        }
    }
    $workbook->close();
    close(BMP);
    close(DEBUG);
}

sub bmp_256c_8bit{#8位图
    say 'pixel deep is 8bit';
    my ($file,$offeset_value,$color_size,$width,$height)=@_;#文件名 最大偏移值 头文件值 图像宽 图像高
    my ($actual_file_name)=($file)=~/(\w+)\./;
    my $real_width=$width%4 ne 0?(4-$width%4 +$width):($width);#bmp 是每行存储的字节数是4的倍数，如果不是4的倍数需要补0
    my $workbook=Excel::Writer::XLSX->new('./'.$actual_file_name.'.xlsx');
    my $worksheet=$workbook->add_worksheet('sheet1');
    $worksheet->set_column(0,$width,2);#把excel的单元格弄成正方形，因为像素是方的    
    open(BMP,$file)||die "error can not open $file,$!\n";
    seek(BMP,$color_size,0)||die "error can seef $file,$!\n";
    my @color_index;#颜色索引
    my $buff;
    foreach(1..($offeset_value-$color_size)/4){
        read(BMP,$buff,4);
        $buff=unpack("H*",$buff);
        $buff=substr(reverse_order($buff,2),2,6);
        push(@color_index,$buff);
    }
    for(1..$height){
        my $height_count=$_;
        for(1..$width){
            my $width_count=$_;
            read(BMP,$buff,1);
            $buff=$color_index[big_hex2dec(unpack("H*",$buff))];
            my $format=$workbook->add_format();
            $format->set_bg_color('#'.$buff);
            $worksheet->write($height-$height_count,$width_count,' ',$format);
        }
        if($real_width-$width>2){
            read(BMP,$buff,(4-$real_width+$width));       
        }else{
            read(BMP,$buff,($real_width-$width));
        }
    }
    $workbook->close();
    close(BMP);
}
sub bmp_16c_4bit{#4bit图
    say 'pixel deep is 4bit';
    my ($file,$offeset_value,$color_size,$width,$height)=@_;#文件名 最大偏移值 头文件值 图像宽 图像高
    my ($actual_file_name)=($file)=~/(\w+)\./;
    my $real_width=$width%4 ne 0?(4-$width%4 +$width):($width);#bmp 是每行存储的字节数是4的倍数，如果不是4的倍数需要补0
    my $workbook=Excel::Writer::XLSX->new('./'.$actual_file_name.'.xlsx');
    my $worksheet=$workbook->add_worksheet('sheet1');   
    $worksheet->set_column(0,$width,1.5);#把excel的单元格弄成正方形，因为像素是方的    
    open(BMP,$file)||die "error can not open $file,$!\n";
    seek(BMP,$color_size,0)||die "error can seef $file,$!\n";
    my $buf_size=$color_size;
    my @color_index;#颜色索引
    my $buff;
    foreach(1..($offeset_value-$color_size)/4){
        read(BMP,$buff,4);
        $buf_size+=4;
        $buff=unpack("H*",$buff);
        $buff=substr(reverse_order($buff,2),2,6);
        push(@color_index,$buff);
    }
    my @stat=stat($file);
    my @color_list;
    for(1..($stat[7]-$buf_size)){
        read(BMP,$buff,1);
        $buff=unpack('H*',$buff);
        push(@color_list,big_hex2dec(substr($buff,0,1)));
        push(@color_list,big_hex2dec(substr($buff,1,1)));
    }
    my $arr_1=0;
    for(1..$height){
        my $height_count=$_;
        for(1..$width){
            my $width_count=$_;
            my $format=$workbook->add_format();
            $format->set_bg_color('#'.$color_index[$color_list[$arr_1++]]);
            $worksheet->write($height-$height_count,$width_count,' ',$format);
        }
        if($real_width-$width<=2){
            $arr_1=$real_width+4-$width+$arr_1;
        }else{
            $arr_1=$real_width+-$width+$arr_1;
        }
        
    }
    $workbook->close();
    close(BMP);
}
sub bmp_1bit_2c{
    say 'pixel deep is 1bit';
    my ($file,$offeset_value,$color_size,$width,$height)=@_;#文件名 最大偏移值 头文件值 图像宽 图像高
    my ($actual_file_name)=($file)=~/(\w+)\./;
    my $real_width=$width%4 ne 0?(4-$width%4 +$width):($width);#bmp 是每行存储的字节数是4的倍数，如果不是4的倍数需要补0
    say 'real_width:'.$real_width;
    say 'width:'.$width;
    my $workbook=Excel::Writer::XLSX->new('./'.$actual_file_name.'.xlsx');
    my $worksheet=$workbook->add_worksheet('sheet1');   
    $worksheet->set_column(0,$width,1.5);#把excel的单元格弄成正方形，因为像素是方的    
    open(BMP,$file)||die "error can not open $file,$!\n";
    seek(BMP,$color_size,0)||die "error can seef $file,$!\n";
    my $buf_size=$color_size;
    my @color_index;#颜色索引
    my $buff;
    foreach(1..($offeset_value-$color_size)/4){
        read(BMP,$buff,4);
        $buf_size+=4;
        $buff=unpack("H*",$buff);
        $buff=substr(reverse_order($buff,2),2,6);
        push(@color_index,$buff);
    }
    my @stat=stat($file);
    my @color_list;
    for(1..($stat[7]-$buf_size)){
        read(BMP,$buff,1);
        $buff=unpack('H*',$buff);
        my $a=sprintf("%b",big_hex2dec($buff));#1bitbmp每个字节存储8个像素的索引值
        $a=length($a)<8?(substr((10**(8-length($a))),1,(8-length($a))).$a):($a);
        my @tep_arr=split(//,$a);
        push(@color_list,@tep_arr);
    }
    my $arr_1=0;
    for(1..$height){
        my $height_count=$_;
        for(1..$width){
            my $width_count=$_;
            my $format=$workbook->add_format();
            $format->set_bg_color('#'.$color_index[$color_list[$arr_1++]]);
            $worksheet->write($height-$height_count,$width_count,' ',$format);
        }
        $arr_1+=($real_width-$width);
    }
    $workbook->close();
    close(BMP);    
}
open(BMP,'./'.$ARGV[0])||die "error can not open $ARGV[0],$!\n";
binmode(BMP);
my $offset_size=0;
read(BMP,my $buf,2);#开始读取文件
if(uc(unpack("H*", $buf)) ne '424D'){
    die "error not a bmp format";
}else{
    say 'format check OK,is bmp format';
}
$offset_size+=2;    
read(BMP,$buf,4);#6
$offset_size+=4;
my $file_size=unpack('H*',$buf);#文件大小
say 'file size is '.(big_hex2dec(reverse_order($file_size,2))/1024/1024).' mb';
read(BMP,$buf,4);#4个保留字段 10
$offset_size+=4;
read(BMP,$buf,4);#14
$offset_size+=4;
my $offset=unpack("H*",$buf);#偏移值
$offset=big_hex2dec(reverse_order($offset,2));
say 'offset valie is '.$offset;
read(BMP,$buf,4);#18
$offset_size+=4;
my $head_size=big_hex2dec(reverse_order(unpack("H*",$buf),2));
say 'head size is :'.$head_size;
read(BMP,$buf,4);#22
$offset_size+=4;
my $width=unpack("H*",$buf);#图像宽
$width=big_hex2dec(reverse_order($width,2));
say 'wdith is '.$width;
read(BMP,$buf,4);#26
$offset_size+=4;
my $height=unpack("H*",$buf);#图像高
$height=big_hex2dec(reverse_order($height,2));
say 'height is '.$height;
read(BMP,$buf,2);#28
$offset_size+=2;
read(BMP,$buf,2);#30
$offset_size+=2;
my $pixel=unpack("H*",$buf);
$pixel=big_hex2dec(reverse_order($pixel,2));
say 'pixel is '.$pixel;
read(BMP,$buf,4);#34
$offset_size+=4;
read(BMP,$buf,4);#水平分辨率 38
$offset_size+=4;
read(BMP,$buf,4);#垂直分辨率 42
$offset_size+=4;
read(BMP,$buf,4);#46
$offset_size+=4;
read(BMP,$buf,4);#50
$offset_size+=4;
read(BMP,$buf,4);#24位bmp 没有调色板数据，文件头一共14+40共54个字节到这边结束了
$offset_size+=4;
say 'actual offset value is '.$offset_size;
close(BMP);
say $pixel;
SWITCH:{
    $pixel==24 && do{bmp_24_bit($ARGV[0],$offset,$width,$height);last SWITCH};
    $pixel==8 && do{bmp_256c_8bit($ARGV[0],$offset,($head_size+14),$width,$height);last SWITCH};
    $pixel==4 && do{bmp_1bit_2c($ARGV[0],$offset,($head_size+14),$width,$height);last SWITCH};
    $pixel==1 && do{bmp_24_bit($ARGV[0],$offset,$width,$height);last SWITCH};
    die "error pixel error!\n";
}
#bmp_24_bit($ARGV[0],$offset,$width,$height);
#bmp_256c_8bit($ARGV[0],$offset,($head_size+14),$width,$height);
#bmp_1bit_2c($ARGV[0],$offset,($head_size+14),$width,$height);
#bmp_16c_4bit($ARGV[0],$offset,($head_size+14),$width,$height);
__END__
=pod

=head1 Vesion

    V0.9build2 20171208 by stefaniexyy

=head1 Description

    Swich a picture to a microsoft excel using pixel

    bmp.pl xxx.bmp and get result xxx.xlsx

=head1 Update list

    20171130 now support bmp 24 bit ture color
    20171208 now support all windows bmp format