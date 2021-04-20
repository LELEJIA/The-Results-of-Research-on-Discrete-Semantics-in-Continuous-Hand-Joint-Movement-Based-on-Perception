library(readxl)
library(irr)
library(ggplot2)
library(ggthemes)
library(xlsx)


# Get Data
# If you want to use this program, please replace the path shown below with the path of the excel file "Statistic Data for the R Program" at the same level as this file.
DataPath1="~/Desktop/LeapMeasure.xlsx"
DataPath2="~/Desktop/LeapMeasureByMotionUnit.xlsx"

Participant1=read_excel(DataPath1,sheet = "ll")
Participant2=read_excel(DataPath1,sheet = "KG")
Participant3=read_excel(DataPath1,sheet = "TS")
Participant4=read_excel(DataPath1,sheet = "JG")
Participant5=read_excel(DataPath1,sheet = "XB")
Participant6=read_excel(DataPath1,sheet = "CY")
Participant7=read_excel(DataPath1,sheet = "ZYQ")
Participant8=read_excel(DataPath1,sheet = "LHL")
Participant9=read_excel(DataPath1,sheet = "WMR")
Participant10=read_excel(DataPath1,sheet = "JY")

ALL_Measured_Errors=read_excel(DataPath2,sheet = "ALL")

# Data Extraction
FMCP_Measure_value=t(cbind(Participant1[,1],Participant2[,1],Participant3[,1],Participant4[,1],Participant5[,1],Participant6[,1],Participant7[,1],Participant8[,1],Participant9[,1],Participant10[,1]))
FPIP_Measure_value=t(cbind(Participant1[,2],Participant2[,2],Participant3[,2],Participant4[,2],Participant5[,2],Participant6[,2],Participant7[,2],Participant8[,2],Participant9[,2],Participant10[,2]))
TMCP_Measure_value=t(cbind(Participant1[,3],Participant2[,3],Participant3[,3],Participant4[,3],Participant5[,3],Participant6[,3],Participant7[,3],Participant8[,3],Participant9[,3],Participant10[,3]))
TIP_Measure_value=t(cbind(Participant1[,4],Participant2[,4],Participant3[,4],Participant4[,4],Participant5[,4],Participant6[,4],Participant7[,4],Participant8[,4],Participant9[,4],Participant10[,4]))
FE_Measure_value=t(cbind(Participant1[,5],Participant2[,5],Participant3[,5],Participant4[,5],Participant5[,5],Participant6[,5],Participant7[,5],Participant8[,5],Participant9[,5],Participant10[,5]))
TE_Measure_value=t(cbind(Participant1[,6],Participant2[,6],Participant3[,6],Participant4[,6],Participant5[,6],Participant6[,6],Participant7[,6],Participant8[,6],Participant9[,6],Participant10[,6]))
WFL_Measure_value=t(cbind(Participant1[,7],Participant2[,7],Participant3[,7],Participant4[,7],Participant5[,7],Participant6[,7],Participant7[,7],Participant8[,7],Participant9[,7],Participant10[,7]))
WEX_Measure_value=t(cbind(Participant1[,8],Participant2[,8],Participant3[,8],Participant4[,8],Participant5[,8],Participant6[,8],Participant7[,8],Participant8[,8],Participant9[,8],Participant10[,8]))
WU_Measure_value=t(cbind(Participant1[,9],Participant2[,9],Participant3[,9],Participant4[,9],Participant5[,9],Participant6[,9],Participant7[,9],Participant8[,9],Participant9[,9],Participant10[,9]))

#Data Save
###
#write.xlsx(abs(FMCP_Mean-FMCP_Standard)/90, DataPath2, row.names = FALSE,append = TRUE,sheetName = "FMCP")
#write.xlsx(abs(FPIP_Mean-FPIP_Standard)/110, DataPath2, row.names = FALSE,append = TRUE,sheetName = "FPIP")
#write.xlsx(abs(TMCP_Mean-TMCP_Standard)/50, DataPath2, row.names = FALSE,append = TRUE,sheetName = "TMCP")
#write.xlsx(abs(TIP_Mean-TIP_Standard)/80, DataPath2, row.names = FALSE,append = TRUE,sheetName = "TIP")
#write.xlsx(abs(-FE_Mean-FE_Standard)/60, DataPath2, row.names = FALSE,append = TRUE,sheetName = "FE")
#write.xlsx(abs(TE_Mean-TE_Standard)/70, DataPath2, row.names = FALSE,append = TRUE,sheetName = "TE")
#write.xlsx(abs(WFL_Mean-WFL_Standard)/80, DataPath2, row.names = FALSE,append = TRUE,sheetName = "WFL")
#write.xlsx(abs(WEX_Mean-WEX_Standard)/70, DataPath2, row.names = FALSE,append = TRUE,sheetName = "WEX")
#write.xlsx(abs(WU_Mean-WU_Standard)/30, DataPath2, row.names = FALSE,append = TRUE,sheetName = "WU")
###

#Mean
FMCP_Mean=apply(FMCP_Measure_value, MARGIN = 2, mean)
FPIP_Mean=apply(FPIP_Measure_value, MARGIN = 2, mean)
TMCP_Mean=apply(TMCP_Measure_value, MARGIN = 2, mean)
TIP_Mean=apply(TIP_Measure_value, MARGIN = 2, mean)
FE_Mean=apply(FE_Measure_value, MARGIN = 2, mean)
TE_Mean=apply(TE_Measure_value, MARGIN = 2, mean)
WFL_Mean=apply(WFL_Measure_value, MARGIN = 2, mean)
WEX_Mean=apply(WEX_Measure_value, MARGIN = 2, mean)
WU_Mean=apply(WU_Measure_value, MARGIN = 2, mean)

#Standard deviation
FMCP_Std=apply(FMCP_Measure_value, MARGIN = 2, sd)
FPIP_Std=apply(FPIP_Measure_value, MARGIN = 2, sd)
TMCP_Std=apply(TMCP_Measure_value, MARGIN = 2, sd)
TIP_Std=apply(TIP_Measure_value, MARGIN = 2, sd)
FE_Std=apply(FE_Measure_value, MARGIN = 2, sd)
TE_Std=apply(TE_Measure_value, MARGIN = 2, sd)
WFL_Std=apply(WFL_Measure_value, MARGIN = 2, sd)
WEX_Std=apply(WEX_Measure_value, MARGIN = 2, sd)
WU_Std=apply(WU_Measure_value, MARGIN = 2, sd)

#
Mean_Charactor=as.character(format(cbind(FMCP_Mean,FPIP_Mean,TMCP_Mean,TIP_Mean,FE_Mean,TE_Mean,WFL_Mean,WEX_Mean,WU_Mean),digits=2))
Std_Charactor=as.character(format(cbind(FMCP_Std,FPIP_Std,TMCP_Std,TIP_Std,FE_Std,TE_Std,WFL_Std,WEX_Std,WU_Std),digits=2))
Mean_std=paste(Mean_Charactor,"(",Std_Charactor,")")
write.xlsx(Mean_std, "~/Desktop/Mean_Std.xlsx", row.names = FALSE,append = TRUE,sheetName = "Mean(std)")

#Standard Angle Value
FMCP_Standard=c(0,22.5,45,67.5,90)
FPIP_Standard=c(0,27.5,55,82.5,110)
TMCP_Standard=c(0,12.5,25,37.5,50)
TIP_Standard=c(0,20,40,60,80)
FE_Standard=c(0,15,30,45,60)
TE_Standard=c(0,17.5,35,52.5,70)
WFL_Standard=c(0,20,40,60,80)
WEX_Standard=c(0,17.5,35,52.5,70)
WU_Standard=c(0,7.5,15,22.5,30)


#Measure Error

The_calibration_state=c(0,0.25,0.5,0.75,1)
percentx=c("0%","25%","50%","75%","100%")
percenty=c("0%","25%","50%","75%","100%")

ggplot(data.frame(FMCP_Standard,FMCP_Mean,The_calibration_state),aes(x=FMCP_Standard,y=abs(FMCP_Mean-FMCP_Standard)/90)) +theme_bw()+ xlab("The Calibration Value (°)") + ylab("Absolute value of error(°)")+ geom_point(col="#3E647D",alpha=0.9,size=3,shape=21,fill="#CCCCCC") +ggtitle("FMCP")+theme(plot.title = element_text(vjust = -7))+theme(plot.title = element_text(hjust = 0.5,size = 14))+ theme(axis.text = element_text(size=14),axis.title=element_text(size=16))
ggplot(data.frame(FPIP_Standard,FPIP_Mean,The_calibration_state),aes(x=FPIP_Standard,y=abs(FPIP_Mean-FPIP_Standard)/110)) +theme_bw()+ xlab("The Calibration Value (°)") + ylab("Absolute value of error(°)")+ geom_point(col="#3E647D",alpha=0.9,size=3,shape=21,fill="#CCCCCC") +ggtitle("FPIP")+theme(plot.title = element_text(vjust = -7))+theme(plot.title = element_text(hjust = 0.5,size = 14))+ theme(axis.text = element_text(size=14),axis.title=element_text(size=16))
ggplot(data.frame(TMCP_Standard,TMCP_Mean,The_calibration_state),aes(x=TMCP_Standard,y=abs(TMCP_Mean-TMCP_Standard)/50)) +theme_bw()+ xlab("The Calibration Value (°)") + ylab("Absolute value of error(°)")+ geom_point(col="#3E647D",alpha=0.9,size=3,shape=21,fill="#CCCCCC") +ggtitle("TMCP")+theme(plot.title = element_text(vjust = -7))+theme(plot.title = element_text(hjust = 0.5,size = 14))+ theme(axis.text = element_text(size=14),axis.title=element_text(size=16))
ggplot(data.frame(TIP_Standard,TIP_Mean,The_calibration_state),aes(x=TIP_Standard,y=abs(TIP_Mean-TIP_Standard)/80)) +theme_bw()+ xlab("The Calibration Value (°)") + ylab("Absolute value of error(°)")+ geom_point(col="#3E647D",alpha=0.9,size=3,shape=21,fill="#CCCCCC") +ggtitle("TIP")+theme(plot.title = element_text(vjust = -7))+theme(plot.title = element_text(hjust = 0.5,size = 14))+ theme(axis.text = element_text(size=14),axis.title=element_text(size=16))
ggplot(data.frame(FE_Standard,FE_Mean,The_calibration_state),aes(x=FE_Standard,y=abs(-FE_Mean-FE_Standard)/60)) +theme_bw()+ xlab("The Calibration Value (°)") + ylab("Absolute value of error(°)")+ geom_point(col="#3E647D",alpha=0.9,size=3,shape=21,fill="#CCCCCC") +ggtitle("FE")+theme(plot.title = element_text(vjust = -7))+theme(plot.title = element_text(hjust = 0.5,size = 14))+ theme(axis.text = element_text(size=14),axis.title=element_text(size=16))
ggplot(data.frame(TE_Standard,TE_Mean,The_calibration_state),aes(x=TE_Standard,y=abs(TE_Mean-TE_Standard)/70)) +theme_bw()+ xlab("The Calibration Value (°)") + ylab("Absolute value of error(°)")+ geom_point(col="#3E647D",alpha=0.9,size=3,shape=21,fill="#CCCCCC") +ggtitle("TE")+theme(plot.title = element_text(vjust = -7))+theme(plot.title = element_text(hjust = 0.5,size = 14))+ theme(axis.text = element_text(size=14),axis.title=element_text(size=16))
ggplot(data.frame(WFL_Standard,WFL_Mean,The_calibration_state),aes(x=WFL_Standard,y=abs(WFL_Mean-WFL_Standard)/80)) +theme_bw()+ xlab("The Calibration Value (°)") + ylab("Absolute value of error(°)")+ geom_point(col="#3E647D",alpha=0.9,size=3,shape=21,fill="#CCCCCC") +ggtitle("WFL")+theme(plot.title = element_text(vjust = -7))+theme(plot.title = element_text(hjust = 0.5,size = 14))+ theme(axis.text = element_text(size=14),axis.title=element_text(size=16))
ggplot(data.frame(WEX_Standard,WEX_Mean,The_calibration_state),aes(x=WEX_Standard,y=abs(WEX_Mean-WEX_Standard)/70)) +theme_bw()+ xlab("The Calibration Value (°)") + ylab("Absolute value of error(°)")+ geom_point(col="#3E647D",alpha=0.9,size=3,shape=21,fill="#CCCCCC") +ggtitle("WEX")+theme(plot.title = element_text(vjust = -7))+theme(plot.title = element_text(hjust = 0.5,size = 14))+ theme(axis.text = element_text(size=14),axis.title=element_text(size=16))
ggplot(data.frame(WU_Standard,WU_Mean,The_calibration_state),aes(x=WU_Standard,y=abs(WU_Mean-WU_Standard)/30)) +theme_bw()+ xlab("The Calibration Value (°)") + ylab("Absolute value of error(°)")+ geom_point(col="#3E647D",alpha=0.9,size=3,shape=21,fill="#CCCCCC") +ggtitle("WU")+theme(plot.title = element_text(vjust = -7))+theme(plot.title = element_text(hjust = 0.5,size = 14))+ theme(axis.text = element_text(size=14),axis.title=element_text(size=16))

#Measure Error ALL
Errors=as.matrix(ALL_Measured_Errors[,1])
Motion_Units=as.matrix(ALL_Measured_Errors[,2])
Calibration_State=as.matrix(ALL_Measured_Errors[,3])
ggplot(data.frame(Errors,Calibration_State,Motion_Units), aes(x=Calibration_State,y=Errors,color=Motion_Units))+geom_point(size = 3.0, shape = 19)+xlab("Requirements Angle(Percentage of Motion Range)") + ylab("Normalized Measured Error")+theme_bw()+theme(axis.text = element_text(size=14),axis.title=element_text(size=16),legend.title = element_text(size = 12), legend.text = element_text(size = 12))+scale_color_brewer("Motion Primitive Units",palette="PuBu")+ expand_limits(y=c(0,1))+scale_y_continuous(labels =percenty)+scale_x_continuous(labels =percentx)
ggsave("all.jpg",width = 8,height = 5,dpi = 500)
#State Dividing Actual
FMCP_Actual=c(0,22.4,70.4,90)
FPIP_Actual=c(0,30.1,80.4,110)
TMCP_Actual=c(0,24.3,50,0)
TIP_Actual=c(0,29.4,80,0)
FE_Actual=c(0,23.4,60,0)
TE_Actual=c(0,27.0,70,0)
WFL_Actual=c(37.2,80,0,0)
WEX_Actual=c(34.1,70,0,0)
WU_Actual=c(-20,11.1,40,0)

#Linear fitting
#FMCP
data.lm<-lm(FMCP_Mean~FMCP_Standard)
formula <- sprintf("y=%.2f%+.2f*x", round(coef(data.lm)[1],4),round(coef(data.lm)[2],4)) 
r2 <- sprintf("R^2= %.2f",summary(data.lm)$r.squared)
FMCP_Leap=(coef(data.lm)[1])+(coef(data.lm)[2])*FMCP_Actual
ggplot(data.frame(FMCP_Standard,FMCP_Mean),aes(x=FMCP_Standard,y=FMCP_Mean)) +theme_few()+ xlab("Actual Motion Value (°)") + ylab("Leap Motion Measured Value (°)")+geom_smooth(method = lm,fill="#CCCCCC",colour="#3E647D")+ geom_point(col="#6794D9",alpha=0.55,size=3,shape=17) +ggtitle(paste(formula,",",r2))+theme(plot.title = element_text(vjust = -12))+theme(plot.title = element_text(hjust = 0.5,size = 14))+ theme(axis.text = element_text(size=14),axis.title=element_text(size=16))
ggsave("1.png",width = 4,height = 4.5,dpi = 300)
#Formula output
FMCP_Formula=formula

#FPIP
data.lm<-lm(FPIP_Mean~FPIP_Standard)
formula <- sprintf("y=%.2f%+.2f*x", round(coef(data.lm)[1],4),round(coef(data.lm)[2],4))
r2 <- sprintf("R^2= %.2f",summary(data.lm)$r.squared)
FPIP_Leap=(coef(data.lm)[1])+(coef(data.lm)[2])*FPIP_Actual
ggplot(data.frame(FPIP_Standard,FPIP_Mean),aes(x=FPIP_Standard,y=FPIP_Mean)) +theme_few()+ xlab("Actual Motion Value (°)") + ylab("Leap Motion Measured Value (°)")+geom_smooth(method = lm,fill="#CCCCCC",colour="#3E647D")+ geom_point(col="#6794D9",alpha=0.55,size=3,shape=17) +ggtitle(paste(formula,",",r2))+theme(plot.title = element_text(vjust = -12))+theme(plot.title = element_text(hjust = 0.5,size = 14))+ theme(axis.text = element_text(size=14),axis.title=element_text(size=16))
ggsave("2.png",width = 4,height = 4.5,dpi = 300)
FPIP_Formula=formula

#TMCP
data.lm<-lm(TMCP_Mean~TMCP_Standard)
formula <- sprintf("y=%.2f%+.2f*x", round(coef(data.lm)[1],4),round(coef(data.lm)[2],4)) 
r2 <- sprintf("R^2= %.2f",summary(data.lm)$r.squared)
TMCP_Leap=(coef(data.lm)[1])+(coef(data.lm)[2])*TMCP_Actual
ggplot(data.frame(TMCP_Standard,TMCP_Mean),aes(x=TMCP_Standard,y=TMCP_Mean)) +theme_few()+ xlab("Actual Motion Value (°)") + ylab("Leap Motion Measured Value (°)")+geom_smooth(method = lm,fill="#CCCCCC",colour="#3E647D")+ geom_point(col="#6794D9",alpha=0.55,size=3,shape=17) +ggtitle(paste(formula,",",r2))+theme(plot.title = element_text(vjust = -12))+theme(plot.title = element_text(hjust = 0.5,size = 14))+ theme(axis.text = element_text(size=14),axis.title=element_text(size=16))
ggsave("3.png",width = 4,height = 4.5,dpi = 300)
TMCP_Formula=formula

#TIP
data.lm<-lm(TIP_Mean~TIP_Standard)
formula <- sprintf("y=%.2f%+.2f*x", round(coef(data.lm)[1],4),round(coef(data.lm)[2],4)) 
r2 <- sprintf("R^2= %.2f",summary(data.lm)$r.squared)
TIP_Leap=(coef(data.lm)[1])+(coef(data.lm)[2])*TIP_Actual
ggplot(data.frame(TIP_Standard,TIP_Mean),aes(x=TIP_Standard,y=TIP_Mean)) +theme_few()+ xlab("Actual Motion Value (°)") + ylab("Leap Motion Measured Value (°)")+geom_smooth(method = lm,fill="#CCCCCC",colour="#3E647D")+ geom_point(col="#6794D9",alpha=0.55,size=3,shape=17) +ggtitle(paste(formula,",",r2))+theme(plot.title = element_text(vjust = -12))+theme(plot.title = element_text(hjust = 0.5,size = 14))+ theme(axis.text = element_text(size=14),axis.title=element_text(size=16))
ggsave("4.png",width = 4,height = 4.5,dpi = 300)
TIP_Formula=formula

#FE
data.lm<-lm(-FE_Mean~FE_Standard)
formula <- sprintf("y=%.2f%+.2f*x", round(coef(data.lm)[1],4),round(coef(data.lm)[2],4)) 
r2 <- sprintf("R^2= %.2f",summary(data.lm)$r.squared)
FE_Leap=(coef(data.lm)[1])+(coef(data.lm)[2])*FE_Actual
ggplot(data.frame(FE_Standard,FE_Mean),aes(x=FE_Standard,y=-FE_Mean)) +theme_few()+ xlab("Actual Motion Value (°)") + ylab("Leap Motion Measured Value (°)")+geom_smooth(method = lm,fill="#CCCCCC",colour="#3E647D")+ geom_point(col="#6794D9",alpha=0.55,size=3,shape=17) +ggtitle(paste(formula,",",r2))+theme(plot.title = element_text(vjust = -12))+theme(plot.title = element_text(hjust = 0.5,size = 14))+ theme(axis.text = element_text(size=14),axis.title=element_text(size=16))
ggsave("5.png",width = 4,height = 4.5,dpi = 300)
FE_Formula=formula

#TE
data.lm<-lm(TE_Mean~TE_Standard)
formula <- sprintf("y=%.2f%+.2f*x", round(coef(data.lm)[1],4),round(coef(data.lm)[2],4)) 
r2 <- sprintf("R^2= %.2f",summary(data.lm)$r.squared)
TE_Leap=(coef(data.lm)[1])+(coef(data.lm)[2])*TE_Actual
ggplot(data.frame(TE_Standard,TE_Mean),aes(x=TE_Standard,y=TE_Mean)) +theme_few()+ xlab("Actual Motion Value (°)") + ylab("Leap Motion Measured Value (°)")+geom_smooth(method = lm,fill="#CCCCCC",colour="#3E647D")+ geom_point(col="#6794D9",alpha=0.55,size=3,shape=17) +ggtitle(paste(formula,",",r2))+theme(plot.title = element_text(vjust = -12))+theme(plot.title = element_text(hjust = 0.5,size = 14))+ theme(axis.text = element_text(size=14),axis.title=element_text(size=16))
ggsave("6.png",width = 4,height = 4.5,dpi = 300)
TE_Formula=formula

#WFL
data.lm<-lm(WFL_Mean~WFL_Standard)
formula <- sprintf("y=%.2f%+.2f*x", round(coef(data.lm)[1],4),round(coef(data.lm)[2],4)) 
r2 <- sprintf("R^2= %.2f",summary(data.lm)$r.squared)
WFL_Leap=(coef(data.lm)[1])+(coef(data.lm)[2])*WFL_Actual
ggplot(data.frame(WFL_Standard,WFL_Mean),aes(x=WFL_Standard,y=WFL_Mean)) +theme_few()+ xlab("Actual Motion Value (°)") + ylab("Leap Motion Measured Value (°)")+geom_smooth(method = lm,fill="#CCCCCC",colour="#3E647D")+ geom_point(col="#6794D9",alpha=0.55,size=3,shape=17) +ggtitle(paste(formula,",",r2))+theme(plot.title = element_text(vjust = -12))+theme(plot.title = element_text(hjust = 0.5,size = 14))+ theme(axis.text = element_text(size=14),axis.title=element_text(size=16))
ggsave("7.png",width = 4,height = 4.5,dpi = 300)
WFL_Formula=formula

#WEX
data.lm<-lm(WEX_Mean~WEX_Standard)
formula <- sprintf("y=%.2f%+.2f*x", round(coef(data.lm)[1],4),round(coef(data.lm)[2],4)) 
r2 <- sprintf("R^2= %.2f",summary(data.lm)$r.squared)
WEX_Leap=(coef(data.lm)[1])+(coef(data.lm)[2])*WEX_Actual
ggplot(data.frame(WEX_Standard,WEX_Mean),aes(x=WEX_Standard,y=WEX_Mean)) +theme_few()+ xlab("Actual Motion Value (°)") + ylab("Leap Motion Measured Value (°)")+geom_smooth(method = lm,fill="#CCCCCC",colour="#3E647D")+ geom_point(col="#6794D9",alpha=0.55,size=3,shape=17) +ggtitle(paste(formula,",",r2))+theme(plot.title = element_text(vjust = -12))+theme(plot.title = element_text(hjust = 0.5,size = 14))+ theme(axis.text = element_text(size=14),axis.title=element_text(size=16))
ggsave("8.png",width = 4,height = 4.5,dpi = 300)
WEX_Formula=formula

#WU
data.lm<-lm(WU_Mean~WU_Standard)
formula <- sprintf("y=%.2f%+.2f*x", round(coef(data.lm)[1],4),round(coef(data.lm)[2],4)) 
r2 <- sprintf("R^2= %.2f",summary(data.lm)$r.squared)
WU_Leap=(coef(data.lm)[1])+(coef(data.lm)[2])*WU_Actual
ggplot(data.frame(WU_Standard,WU_Mean),aes(x=WU_Standard,y=WU_Mean)) +theme_few()+ xlab("Actual Motion Value (°)") + ylab("Leap Motion Measured Value (°)")+geom_smooth(method = lm,fill="#CCCCCC",colour="#3E647D")+ geom_point(col="#6794D9",alpha=0.55,size=3,shape=17) +ggtitle(paste(formula,",",r2))+theme(plot.title = element_text(vjust = -12))+theme(plot.title = element_text(hjust = 0.5,size = 14))+ theme(axis.text = element_text(size=14),axis.title=element_text(size=16))
ggsave("9.png",width = 4,height = 4.5,dpi = 300)
WU_Formula=formula

#ALL Formula
all_formula=cbind(FMCP_Formula,FPIP_Formula,TMCP_Formula,TIP_Formula,FE_Formula,TE_Formula,WFL_Formula,WEX_Formula,WU_Formula)
write.xlsx(all_formula, "~/Desktop/All_Formula.xlsx", row.names = FALSE,append = TRUE,sheetName = "formulas")

#ALL Dividing Values
all_Dividing_Values=cbind(FMCP_Leap,FPIP_Leap,TMCP_Leap,TIP_Leap,FE_Leap,TE_Leap,WFL_Leap,WEX_Leap,WU_Leap)
write.xlsx(all_Dividing_Values, "~/Desktop/all_Dividing_Values.xlsx", row.names = FALSE,append = TRUE,sheetName = "all_Dividing_Values")

#


