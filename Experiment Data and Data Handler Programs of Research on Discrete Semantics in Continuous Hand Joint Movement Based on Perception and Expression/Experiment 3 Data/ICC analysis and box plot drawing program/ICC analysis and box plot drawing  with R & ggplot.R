library(readxl)
library(irr)
library(ggplot2)
library(ggthemes)


# Get Data
# If you want to use this program, please replace the path shown below with the path of the excel file "Statistic Data for the R Program" at the same level as this file.
DataPath="~/Desktop/Experiment Data and Data Handler Program/Experiment 3 Data/ICC analysis and box plot drawing program/Statistic Data for the R Program.xlsx"

FMCP_ICC=read_excel(DataPath,sheet = "Finger MCP")
FPIP_ICC=read_excel(DataPath,sheet = "Finger PIP")
TMCP_ICC=read_excel(DataPath,sheet = "Thumb MCP")
TIP_ICC=read_excel(DataPath,sheet = "Thumb IP")
FE_ICC=read_excel(DataPath,sheet = "Finger Abduction")
TE_ICC=read_excel(DataPath,sheet = "Thumb Abduction")
WFL_ICC=read_excel(DataPath,sheet = "Wrist Flexion")
WEX_ICC=read_excel(DataPath,sheet = "Wrist Extension")
WU_ICC=read_excel(DataPath,sheet = "Wrist Ulnar Deviation")
WR_ICC=read_excel(DataPath,sheet = "Wrist Radial Deviation")

# Separate Data
FMCP_ICC1=FMCP_ICC$`DATA`[seq(1,length(FMCP_ICC$DATA),2)]
FMCP_ICC2=FMCP_ICC$`DATA`[seq(2,length(FMCP_ICC$DATA),2)]

FPIP_ICC1=FPIP_ICC$`DATA`[seq(1,length(FPIP_ICC$DATA),2)]
FPIP_ICC2=FPIP_ICC$`DATA`[seq(2,length(FPIP_ICC$DATA),2)]

TMCP_ICC1=TMCP_ICC$`DATA`[seq(1,length(TMCP_ICC$DATA),2)]
TMCP_ICC2=TMCP_ICC$`DATA`[seq(2,length(TMCP_ICC$DATA),2)]

TIP_ICC1=TIP_ICC$`DATA`[seq(1,length(TIP_ICC$DATA),2)]
TIP_ICC2=TIP_ICC$`DATA`[seq(2,length(TIP_ICC$DATA),2)]

FE_ICC1=FE_ICC$`DATA`[seq(1,length(FE_ICC$DATA),2)]
FE_ICC2=FE_ICC$`DATA`[seq(2,length(FE_ICC$DATA),2)]

TE_ICC1=TE_ICC$`DATA`[seq(1,length(TE_ICC$DATA),2)]
TE_ICC2=TE_ICC$`DATA`[seq(2,length(TE_ICC$DATA),2)]

WFL_ICC1=WFL_ICC$`DATA`[seq(1,length(WFL_ICC$DATA),2)]
WFL_ICC2=WFL_ICC$`DATA`[seq(2,length(WFL_ICC$DATA),2)]

WEX_ICC1=WEX_ICC$`DATA`[seq(1,length(WEX_ICC$DATA),2)]
WEX_ICC2=WEX_ICC$`DATA`[seq(2,length(WEX_ICC$DATA),2)]

WU_ICC1=WU_ICC$`DATA`[seq(1,length(WU_ICC$DATA),2)]
WU_ICC2=WU_ICC$`DATA`[seq(2,length(WU_ICC$DATA),2)]

WR_ICC1=WR_ICC$`DATA`[seq(1,length(WR_ICC$DATA),2)]
WR_ICC2=WR_ICC$`DATA`[seq(2,length(WR_ICC$DATA),2)]

# Calculate ICC
FMCP_ICC_value=icc(cbind(FMCP_ICC1,FMCP_ICC2), model = "twoway", type = "agreement", unit = "single")
FPIP_ICC_value=icc(cbind(FPIP_ICC1,FPIP_ICC2), model = "twoway", type = "agreement", unit = "single")
TMCP_ICC_value=icc(cbind(TMCP_ICC1,TMCP_ICC2), model = "twoway", type = "agreement", unit = "single")
TIP_ICC_value=icc(cbind(TIP_ICC1,TIP_ICC2), model = "twoway", type = "agreement", unit = "single")
FE_ICC_value=icc(cbind(FE_ICC1,FE_ICC2), model = "twoway", type = "agreement", unit = "single")
TE_ICC_value=icc(cbind(TE_ICC1,TE_ICC2), model = "twoway", type = "agreement", unit = "single")
WFL_ICC_value=icc(cbind(WFL_ICC1,WFL_ICC2), model = "twoway", type = "agreement", unit = "single")
WEX_ICC_value=icc(cbind(WEX_ICC1,WEX_ICC2), model = "twoway", type = "agreement", unit = "single")
WU_ICC_value=icc(cbind(WU_ICC1,WU_ICC2), model = "twoway", type = "agreement", unit = "single")
WR_ICC_value=icc(cbind(WR_ICC1,WR_ICC2), model = "twoway", type = "agreement", unit = "single")

ICC_VALUE<-c(FMCP_ICC_value$value,FPIP_ICC_value$value,TMCP_ICC_value$value,TIP_ICC_value$value,FE_ICC_value$value,TE_ICC_value$value,WFL_ICC_value$value,WEX_ICC_value$value,WU_ICC_value$value)
ICC_P<-c(FMCP_ICC_value$p.value,FPIP_ICC_value$p.value,TMCP_ICC_value$p.value,TIP_ICC_value$p.value,FE_ICC_value$p.value,TE_ICC_value$p.value,WFL_ICC_value$p.value,WEX_ICC_value$p.value,WU_ICC_value$p.value)
ICC_VALUE
ICC_P
#Single Average
FMCP=(FMCP_ICC1+FMCP_ICC2)/2
FPIP=(FPIP_ICC1+FPIP_ICC2)/2
TMCP=(TMCP_ICC1+TMCP_ICC2)/2
TIP=(TIP_ICC1+TIP_ICC2)/2
FE=(FE_ICC1+FE_ICC2)/2
TE=(TE_ICC1+TE_ICC2)/2
WFL=(WFL_ICC1+WFL_ICC2)/2
WEX=(WEX_ICC1+WEX_ICC2)/2
WU=(WU_ICC1+WU_ICC2)/2
WR=(WR_ICC1+WR_ICC2)/2

#Draw BoxPlot
#FMCP
FMCP_Box=factor(rep(c("State 1","State 2","State 3"),each=length(FMCP)/3))
#boxplot(FMCP~FMCP_Box,data.frame(FMCP,FMCP_Box))
ggplot(data.frame(FMCP,FMCP_Box),aes(FMCP_Box,FMCP))+stat_boxplot(geom = "errorbar",width=0.15)+geom_hline(aes(yintercept=22.4),color="#BFBFBF")+geom_hline(aes(yintercept=70.7),color="#BFBFBF")+geom_boxplot(fill=c("#6794A7", "#014D64", "#01A2D9"))+stat_summary(fun.y=mean, geom="point", position=position_dodge(.9), shape=23, size=4,fill='white')+ theme_bw()+theme(axis.title.x =element_text(size=18), axis.title.y=element_text(size=18),axis.text = element_text(size=18),panel.grid.major =element_blank(), panel.grid.minor = element_blank(),panel.background = element_blank(),axis.line = element_line(colour = "black"))+labs(x="Semantic state", y = "Movement angle (°)")
ggsave("figure 1.jpg",width =4, height =3.5,dpi = 500)
#FPIP
FPIP_Box=factor(rep(c("State 1","State 2","State 3"),each=length(FPIP)/3))
#boxplot(FPIP~FPIP_Box,data.frame(FPIP,FPIP_Box))
ggplot(data.frame(FPIP,FPIP_Box),aes(FPIP_Box,FPIP))+ stat_boxplot(geom = "errorbar",width=0.15)+geom_boxplot(fill=c("#6794A7", "#014D64", "#01A2D9"))+geom_hline(aes(yintercept=30.1),color="#BFBFBF")+geom_hline(aes(yintercept=80.4),color="#BFBFBF")+stat_summary(fun.y=mean, geom="point", position=position_dodge(.9), shape=23, size=4,fill='white')+ theme_bw()+theme(axis.title.x =element_text(size=18), axis.title.y=element_text(size=18),axis.text = element_text(size=18),panel.grid.major =element_blank(), panel.grid.minor = element_blank(),panel.background = element_blank(),axis.line = element_line(colour = "black"))+labs(x="Semantic state", y = "Movement angle (°)")
ggsave("figure 2.jpg",width =4, height =3.5,dpi = 500)
#TMCP
x1=c(0,53)
y1=c("State 1","State 2")
TMCP_Box=factor(rep(c("State 1","State 2"),each=length(TMCP)/2))
#boxplot(TMCP~TMCP_Box,data.frame(TMCP,TMCP_Box))
ggplot(data.frame(TMCP,TMCP_Box),aes(TMCP_Box,TMCP))+ stat_boxplot(geom = "errorbar",width=0.15)+geom_boxplot(fill=c("#6794A7", "#014D64"))+geom_hline(aes(yintercept=37.2),color="#BFBFBF")+theme_bw()+theme(axis.title.x =element_text(size=18), axis.title.y=element_text(size=18),axis.text = element_text(size=18),panel.grid.major =element_blank(), panel.grid.minor = element_blank(),panel.background = element_blank(),axis.line = element_line(colour = "black"))+labs(x="Semantic state", y = "Movement angle (°)")+geom_point(aes(x="State 1",y=5), geom="point", position=position_dodge(.9), shape=23, size=4,fill='white')+geom_point(aes(x="State 2",y=49), geom="point", position=position_dodge(.9), shape=23, size=4,fill='white')
ggsave("figure 3.jpg",width =4, height =3.5,dpi = 500)
#TIP
TIP_Box=factor(rep(c("State 1","State 2"),each=length(TIP)/2))
#boxplot(TIP~TIP_Box,data.frame(TIP,TIP_Box))
ggplot(data.frame(TIP,TIP_Box),aes(TIP_Box,TIP))+ stat_boxplot(geom = "errorbar",width=0.15)+geom_boxplot(fill=c("#6794A7", "#014D64"))+geom_hline(aes(yintercept=29.4),color="#BFBFBF")+ theme_bw()+theme(axis.title.x =element_text(size=18), axis.title.y=element_text(size=18),axis.text = element_text(size=18),panel.grid.major =element_blank(), panel.grid.minor = element_blank(),panel.background = element_blank(),axis.line = element_line(colour = "black"))+labs(x="Semantic state", y = "Movement angle (°)")+geom_point(aes(x="State 1",y=0), geom="point", position=position_dodge(.9), shape=23, size=4,fill='white')+geom_point(aes(x="State 2",y=74), geom="point", position=position_dodge(.9), shape=23, size=4,fill='white')
ggsave("figure 4.jpg",width =4, height =3.5,dpi = 500)
#FE
FE_Box=factor(rep(c("State 1","State 2"),each=length(FE)/2))
#boxplot(FE~FE_Box,data.frame(FE,FE_Box))
ggplot(data.frame(FE,FE_Box),aes(FE_Box,FE))+ stat_boxplot(geom = "errorbar",width=0.15)+geom_boxplot(fill=c("#6794A7", "#014D64"))+geom_hline(aes(yintercept=23.4),color="#BFBFBF")+ theme_bw()+theme(axis.title.x =element_text(size=18), axis.title.y=element_text(size=18),axis.text = element_text(size=18),panel.grid.major =element_blank(), panel.grid.minor = element_blank(),panel.background = element_blank(),axis.line = element_line(colour = "black"))+labs(x="Semantic state", y = "Movement angle (°)")+geom_point(aes(x="State 1",y=0), geom="point", position=position_dodge(.9), shape=23, size=4,fill='white')+geom_point(aes(x="State 2",y=44), geom="point", position=position_dodge(.9), shape=23, size=4,fill='white')
ggsave("figure 5.jpg",width =4, height =3.5,dpi = 500)
#TE
TE_Box=factor(rep(c("State 1","State 2"),each=length(TE)/2))
#boxplot(TE~TE_Box,data.frame(TE,TE_Box))
ggplot(data.frame(TE,TE_Box),aes(TE_Box,TE))+ stat_boxplot(geom = "errorbar",width=0.15)+geom_boxplot(fill=c("#6794A7", "#014D64"))+geom_hline(aes(yintercept=27),color="#BFBFBF")+ theme_bw()+theme(axis.title.x =element_text(size=18), axis.title.y=element_text(size=18),axis.text = element_text(size=18),panel.grid.major =element_blank(), panel.grid.minor = element_blank(),panel.background = element_blank(),axis.line = element_line(colour = "black"))+labs(x="Semantic state", y = "Movement angle (°)")+geom_point(aes(x="State 1",y=0), geom="point", position=position_dodge(.9), shape=23, size=4,fill='white')+geom_point(aes(x="State 2",y=53), geom="point", position=position_dodge(.9), shape=23, size=4,fill='white')
ggsave("figure 6.jpg",width =4, height =3.5,dpi = 500)
#WFL
WFL_Box=factor(rep(c("State 1","State 2"),each=length(WFL)/2))
#boxplot(WFL~WFL_Box,data.frame(WFL,WFL_Box))
ggplot(data.frame(WFL,WFL_Box),aes(WFL_Box,WFL))+ stat_boxplot(geom = "errorbar",width=0.15)+geom_boxplot(fill=c("#6794A7", "#014D64"))+geom_hline(aes(yintercept=37.2),color="#BFBFBF")+stat_summary(fun.y=mean, geom="point", position=position_dodge(.9), shape=23, size=4,fill='white')+ theme_bw()+theme(axis.title.x =element_text(size=18), axis.title.y=element_text(size=18),axis.text = element_text(size=18),panel.grid.major =element_blank(), panel.grid.minor = element_blank(),panel.background = element_blank(),axis.line = element_line(colour = "black"))+labs(x="Semantic state", y = "Movement angle (°)")
ggsave("figure 7.jpg",width =4, height =3.5,dpi = 500)
#WEX
WEX_Box=factor(rep(c("State 1","State 2"),each=length(WEX)/2))
#boxplot(WEX~WEX_Box,data.frame(WEX,WEX_Box))
ggplot(data.frame(WEX,WEX_Box),aes(WEX_Box,WEX))+ stat_boxplot(geom = "errorbar",width=0.15)+geom_boxplot(fill=c("#6794A7", "#014D64"))+geom_hline(aes(yintercept=34.1),color="#BFBFBF")+stat_summary(fun.y=mean, geom="point", position=position_dodge(.9), shape=23, size=4,fill='white')+ theme_bw()+theme(axis.title.x =element_text(size=18), axis.title.y=element_text(size=18),axis.text = element_text(size=18),panel.grid.major =element_blank(), panel.grid.minor = element_blank(),panel.background = element_blank(),axis.line = element_line(colour = "black"))+labs(x="Semantic state", y = "Movement angle (°)")
ggsave("figure 8.jpg",width =4, height =3.5,dpi = 500)
#WU
WU_Box=factor(rep(c("State 1","State 2"),each=length(WU)/2))
#boxplot(WU~WU_Box,data.frame(WU,WU_Box))
ggplot(data.frame(WU,WU_Box),aes(WU_Box,WU))+ stat_boxplot(geom = "errorbar",width=0.15)+geom_boxplot(fill=c("#6794A7", "#014D64"))+geom_hline(aes(yintercept=11.1),color="#BFBFBF")+stat_summary(fun.y=mean, geom="point", position=position_dodge(.9), shape=23, size=4,fill='white')+ theme_bw()+theme(axis.title.x =element_text(size=18), axis.title.y=element_text(size=18),axis.text = element_text(size=18),panel.grid.major =element_blank(), panel.grid.minor = element_blank(),panel.background = element_blank(),axis.line = element_line(colour = "black"))+labs(x="Semantic state", y = "Movement angle (°)")
ggsave("figure 9.jpg",width =4, height =3.5,dpi = 500)
#WR
WR_Box=factor(rep(c("State 1"),each=length(WU)))
#boxplot(WU~WU_Box,data.frame(WU,WU_Box))
ggplot(data.frame(WR,WR_Box),aes(WR_Box,WR))+ stat_boxplot(geom = "errorbar",width=0.15)+geom_boxplot(fill=c("#6794A7"))+stat_summary(fun.y=mean, geom="point", position=position_dodge(.9), shape=23, size=4,fill='white')+ theme_bw()+theme(axis.title.x =element_text(size=18), axis.title.y=element_text(size=18),axis.text = element_text(size=18),panel.grid.major =element_blank(), panel.grid.minor = element_blank(),panel.background = element_blank(),axis.line = element_line(colour = "black"))+labs(x="Semantic state", y = "Movement angle (°)")
ggsave("figure 10.jpg",width =4, height =3.5,dpi = 500)

