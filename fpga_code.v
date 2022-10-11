module reciever(on_period,baud_1,out_2,clock,serial_2);
output on_period,baud_1;
output [21:0] out_2;
input clock,serial_2;
integer count_1=0;
integer index=0;
reg baud_1=1'b1;
reg on_period=1'b1;
reg [21:0] out_2;
always@(posedge clock)
begin
	if(count_1==1302)
		begin
			baud_1=~baud_1;
			count_1=0;
		end
	else
		count_1=count_1+1;
end
always@(posedge baud_1)
begin
	if(index<29 && on_period==1'b0)
		begin
			out_2[index]=serial_2;
			index=index+1;
		end
	else if(serial_2==1'b0)
		on_period=1'b0;
	else
		begin
			index=0;
			on_period=1'b1;
		end
end
endmodule