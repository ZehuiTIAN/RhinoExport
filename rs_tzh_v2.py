#encoding=utf-8

import rhinoscriptsyntax as rs
from System.Windows.Forms import Form, DialogResult, Label, Button, TextBox
from System.Drawing import Point, Size
import rhinoscript.selection
import rhinoscript.geometry
import re
import sys

reload(sys)

sys.setdefaultencoding('utf-8')

class Joint:
    def __init__(self,number,X,Y,Z):
        self.m_number=number
        self.m_X=float(X)
        self.m_Y=float(Y)
        self.m_Z=float(Z)
        self.m_GlobalX=float(X)
        self.m_GlobalY=float(Y)
        self.m_GlobalZ=float(Z)
        # 初始化时不从外部赋值的属性
        self.m_CoordSys="GLOBAL"
        self.m_CoordType="Cartesian"
        self.m_SpecialJt="No"
        self.m_Frames=[]
    
    #计算该节点出发的杆件间的最大最小夹角
    def cal_max_min_frame_angle(self):
        if self.Frames is None:
            print("No Frames!")
            return -1
        else:
            max_angle=0
            min_angle=180
            max_frame1=self.m_Frames[0]
            max_frame2=self.m_Frames[0]
            min_frame1=self.m_Frames[0]
            min_frame2=self.m_Frames[0]
            for i in range(len(self.m_Frames)-1):
                if self.m_Frames[i].start_joint!=self:
                    vec1=-self.m_Frames[i].vec
                else:
                    vec1=self.m_Frames[i].vec
                for j in range(i+1,len(self.m_Frames)):
                    if self.m_Frames[j].start_joint!=self:
                        vec2=-self.Frames[j].vec
                    else:
                        vec2=self.m_Frames[j].vec
                    tmp_angle=rs.VectorAngle(vec1,vec2)
                    if tmp_angle>max_angle:
                        max_angle=tmp_angle
                        max_frame1=self.m_Frames[i]
                        max_frame2=self.m_Frames[j]
                    if tmp_angle<min_angle:
                        min_angle=tmp_angle
                        min_frame1=self.m_Frames[i]
                        min_frame2=self.m_Frames[j]
        return max_angle,max_frame1,max_frame2,min_angle,min_frame1,min_frame2

class Frame:
    def __init__(self,number,start_joint,end_joint,GUID):
        self.m_number=number
        self.m_start_joint=start_joint
        self.m_end_joint=end_joint
        self.m_vector=rs.CreateVector(end_joint.m_GlobalX-start_joint.m_GlobalX,end_joint.m_GlobalY-start_joint.m_GlobalY,end_joint.m_GlobalZ-start_joint.m_GlobalZ)
        self.m_CentroidX=(start_joint.m_GlobalX+end_joint.m_GlobalX)/2
        self.m_CentroidY=(start_joint.m_GlobalY+end_joint.m_GlobalY)/2
        self.m_CentroidZ=(start_joint.m_GlobalZ+end_joint.m_GlobalZ)/2
        self.m_Length=((start_joint.m_GlobalX+end_joint.m_GlobalX)**2+(start_joint.m_GlobalY+end_joint.m_GlobalY)**2+(start_joint.m_GlobalZ+end_joint.m_GlobalZ)**2)**0.5
        self.m_GUID=GUID
        self.m_IsCurved = "No"

class Area:
    def __init__(self,number,joint_count,joints,GUID):
        self.m_number=number
        self.m_joint_count=joint_count
        self.m_joints=joints
        self.m_GUID=GUID
        
class ModelInfo:
    def __init__(self):
        #构件数据：object data
        self.m_frames=[]
        self.m_areas=[]
        self.m_joints_dict={}
       
        #以下属性表名从文本文件中整行读取，不进行去空格和去换行符处理
        #导出时也直接照读取进来的原样打印不用加换行符
        self.m_s2k_frame_table_names=[]#s2k中每根杆件对应的属性表名
        self.m_s2k_area_table_names=[]#s2k中每根杆件对应的属性表名
        self.m_mgt_obj_table_names=[]#mgt中每根杆件对应的属性表名
        self.m_s2k_docu_table_names=[]#s2k中其他属性表名
        self.m_mgt_docu_table_names=[]#mgt中其他属性表名
        
    
    def set_data(self):
        objectIds = rs.GetObjects("Select")
        if objectIds is None: return
        j_n=0
        f_n=0
        a_n=0
        for objectId in objectIds:
            # 将端点存入字典,键类型为代表坐标的三维元组，值类型为joint
            #如果类型为4（Curve）
            if rs.ObjectType(objectId)==4:
                start_x,start_y,start_z=rs.CurveStartPoint(objectId)
                start_coord=(start_x,start_y,start_z)
                if self.m_joints_dict.has_key(start_coord):
                    start_joint=self.m_joints_dict[start_coord]
                else:
                    j_n+=1
                    start_joint=Joint(j_n,start_x,start_y,start_z)
                    self.m_joints_dict[start_coord]=start_joint
                end_x,end_y,end_z=rs.CurveEndPoint(objectId)
                end_coord=(end_x,end_y,end_z)
                if self.m_joints_dict.has_key(end_coord):
                    end_joint=self.m_joints_dict[end_coord]
                else:
                    j_n+=1
                    end_joint=Joint(j_n,end_x,end_y,end_z)
                    self.m_joints_dict[end_coord]=end_joint
                # 将杆件存入列表
                f_n+=1
                frame=Frame(f_n,start_joint,end_joint,objectId)
                self.m_frames.append(frame)
                self.m_joints_dict[start_coord].m_Frames.append(frame)
                self.m_joints_dict[end_coord].m_Frames.append(frame)
            #如果类型为8（Surface）
            elif rs.ObjectType(objectId)==8:
                area_joints=[]
                points=rs.SurfacePoints(objectId)
                for pt in points:
                    if self.m_joints_dict.has_key(pt):
                        area_joints.append(self.m_joints_dict[pt])
                    else:
                        j_n+=1
                        tmp_j=Joint(j_n,pt[0],pt[1],pt[2])
                        self.m_joints_dict[pt]=tmp_j
                        area_joints.append(tmp_j)
                a_n+=1
                j_c=len(area_joints)
                area=Area(a_n,j_c,area_joints,objectId)
                self.m_areas.append(area)
        #读取documentdata中存下来的表名
        if(rs.GetDocumentData("s2k_frame_table_names")):
            self.m_s2k_frame_table_names=rs.GetDocumentData("s2k_frame_table_names")
        if(rs.GetDocumentData("s2k_area_table_names")):
            self.m_s2k_area_table_names=rs.GetDocumentData("s2k_area_table_names")
        if(rs.GetDocumentData("mgt_obj_table_names")):
            self.m_mgt_obj_table_names=rs.GetDocumentData("mgt_obj_table_names")
        if(rs.GetDocumentData("s2k_docu_table_names")):
            self.m_s2k_docu_table_names=rs.GetDocumentData("s2k_docu_table_names")
        if(rs.GetDocumentData("mgt_docu_table_names")):
            self.m_mgt_docu_table_names=rs.GetDocumentData("mgt_docu_table_names")
        return
     
    def export_s2k(self):
        self.set_data()
        #prompt the user to specify a file name
        filter = "s2k File (*.s2k)|*.s2k|All files (*.*)|*.*||"
        filename = rs.SaveFileName("Save model Points As", filter)
        if not filename: return
    
        #open a new file
        file = open( filename, "w" )
    
        # file path
        file.write("File %s was saved on m/d/yy at h:mm:ss\n"%filename)
        file.write("\n")
        
        for nm in self.m_s2k_docu_table_names:
            if(rs.GetDocumentData(nm)):
                file.write(nm)
                for entry in rs.GetDocumentData(nm):
                    file.write(entry)
                file.write("\n")
                
        #以下两张为特殊表格
        #table JOINT COORDINATES
        file.write("TABLE:  \"JOINT COORDINATES\"\n")
        for coord in self.m_joints_dict:
            joint=self.m_joints_dict[coord]
            file.write("   Joint=%s"%joint.m_number)
            file.write("   CoordSys=%s"%(joint.m_CoordSys))
            file.write("   CoordType=%s"%(joint.m_CoordType))
            file.write("   XorR=%s   Y=%s   Z=%s"%(joint.m_X,joint.m_Y,joint.m_Z))
            file.write("   SpecialJt=%s"%(joint.m_SpecialJt))
            file.write("   GlobalX=%s   GlobalY=%s   GlobalZ=%s"%(joint.m_GlobalX,joint.m_GlobalY,joint.m_GlobalZ))
            file.write("   GUID=")
            file.write("\n")
        file.write("\n")
    
        #table CONNECTIVITY - FRAME
        file.write("TABLE:  \"CONNECTIVITY - FRAME\"\n")
        for frame in self.m_frames:
            file.write("   Frame=%s"%frame.m_number)
            file.write("   JointI=%s   JointJ=%s"%(frame.m_start_joint.m_number,frame.m_end_joint.m_number))
            file.write("   IsCurved=%s"%frame.m_IsCurved)
            file.write("   Length=%s"%frame.m_Length)
            file.write("   CentroidX=%s   CentroidY=%s   CentroidZ=%s"%(frame.m_CentroidX,frame.m_CentroidY,frame.m_CentroidZ))
            file.write("   GUID=%s"%frame.m_GUID)
            file.write("\n")
        file.write("\n")
        
        #table CONNECTIVITY - AREA
        file.write("TABLE:  \"CONNECTIVITY - AREA\"\n")
        for area in self.m_areas:
            file.write("   Area=%s"%area.m_number)
            file.write("   NumJoints=%s"%area.m_joint_count)
            for i in range(area.m_joint_count):
                file.write("   Joint%s=%s"%(i+1,area.m_joints[i].m_number))
            file.write("   GUID=%s"%area.m_GUID)
            file.write("\n")
        file.write("\n")
        
        for nm in self.m_s2k_frame_table_names:
            file.write(nm)
            for fm in self.m_frames:
                if rs.GetUserText(fm.m_GUID,nm) is not None:
                    file.write(rs.GetUserText(fm.m_GUID,nm))
            file.write("\n")
        
        for nm in self.m_s2k_area_table_names:
            file.write(nm)
            for ar in self.m_areas:
                if rs.GetUserText(ar.m_GUID,nm) is not None:
                    file.write(rs.GetUserText(ar.m_GUID,nm))
            file.write("\n")
        
        file.write("END TABLE DATA")
        #close the file
        file.close()
    
    def import_s2k(self):
        #prompt the user for a file to import
        filter = "s2k file (*.s2k)|*.s2k|All Files (*.*)|*.*||"
        filename = rs.OpenFileName("Open Point File", filter)
        if not filename: return
    
        #read each line from the file
        file = open(filename, "r")
        table=""
        joint_coordinates=[]
        connectivity_frame=[]
        connectivity_area=[]
        frame_text={}
        area_text={}
        #frame_section_assignments=[]
        #material_properties_01=[]
        #section_properties_01=[]
        unfinished_line=""
        for line in file:
            whole_line=unfinished_line+line
            ls=whole_line.split()
            #对空行或末尾行或未结束行进行判断
            if len(ls)==0:
                continue
            if (len(ls)>=3) and (ls[0]=="END") and (ls[1]=="TABLE") and (ls[2]=="DATA"):
                break
            if (ls[-1]=="_"):
                unfinished_line=whole_line
                continue
            else:
                unfinished_line=""
            #如果是表名行
            if ls[0]=="TABLE:":
                #三张特殊表格
                #table JOINT COORDINATES
                if (len(ls)>=3) and (ls[1]=="\"JOINT") and (ls[2]=="COORDINATES\""):
                    table="TABLE:  \"JOINT COORDINATES\"\n"
                    continue
                #table CONNECTIVITY - FRAME
                if (len(ls)>=3) and (ls[1]=="\"CONNECTIVITY") and (ls[2]=="-") and (ls[3]=="FRAME\""):
                    table="TABLE:  \"CONNECTIVITY - FRAME\"\n"
                    continue
                #table CONNECTIVITY - AREA
                if (len(ls)>=3) and (ls[1]=="\"CONNECTIVITY") and (ls[2]=="-") and (ls[3]=="FRAME\""):
                    table="TABLE:  \"CONNECTIVITY - AREA\"\n"
                    continue
                #其他表格
                else:
                    table=whole_line
                    continue
            
            #如果不是表名行
            else:
                #不是表格数据，跳过
                if table=="":
                    continue
                    
                #是三张特殊表数据（构件位置数据）
                elif table=="TABLE:  \"JOINT COORDINATES\"\n":
                    joint_coordinates.append(ls)
                elif table=="TABLE:  \"CONNECTIVITY - FRAME\"\n":
                    connectivity_frame.append(ls)
                elif table=="TABLE:  \"CONNECTIVITY - AREA\"\n":
                    connectivity_area.append(ls)
                #是其他表数据，下面判断是document data还是object text
                else:
                    tmp=ls[0].split("=")
                    #是跟随每个Frame的text，存入一个中间结构frame_text
                    if len(tmp)>=2 and tmp[0]=="Frame":
                        if not frame_text.has_key(table):
                            tmp_dic={}
                            frame_text[table]=tmp_dic
                            rs.SetDocumentData("s2k_frame_table_names",str(table)," ")
                        frame_text[table][tmp[1]]=whole_line
                    #是跟随每个Area的text，存入一个中间结构area_text
                    if len(tmp)>=2 and tmp[0]=="Area":
                        if not area_text.has_key(table):
                            tmp_dic={}
                            area_text[table]=tmp_dic
                            rs.SetDocumentData("s2k_area_table_names",str(table)," ")
                        area_text[table][tmp[1]]=whole_line
                            
                    #是document data 直接存
                    else:
                        rs.SetDocumentData("s2k_docu_table_names",str(table)," ")
                        rs.SetDocumentData(table,whole_line," ")
        file.close()
    
        joints_dict={} #以编号为key，以坐标元组为value
        frames_dict={} #以编号为key，以GUID为value
        areas_dict={}  #以编号为key，以GUID为value
        #生成joints_dict
        for i in joint_coordinates:
            number=-1
            X=0.0
            Y=0.0
            Z=0.0
            for j in i:
                l=j.split("=")
                if (len(l)==2) and l[0]=="Joint":
                    number=l[1]
                elif (len(l)==2) and l[0]=="XorR":
                    X=l[1]
                elif (len(l)==2) and l[0]=="Y":
                    Y=l[1]
                elif (len(l)==2) and l[0]=="Z":
                    Z=l[1]
            if number!=-1:
                joints_dict[number]=(X,Y,Z)
        #生成frames_dict
        for i in connectivity_frame:
            jointi=-1
            jointj=-1
            frame_n=-1
            for j in i:
                l=j.split("=")
                if (len(l)==2) and l[0]=="Frame":
                    frame_n=l[1]
                elif (len(l)==2) and l[0]=="JointI":
                    if joints_dict.has_key(l[1]):
                        jointi=joints_dict[l[1]]
                    else:
                        continue
                elif (len(l)==2) and l[0]=="JointJ":
                    if joints_dict.has_key(l[1]):
                        jointj=joints_dict[l[1]]
                    else:
                        continue
            if (jointi!=-1) and (jointj!=-1) and (frame_n!=-1):
                obj=rs.AddLine(jointi,jointj)
                frames_dict[frame_n]=obj
                for tb_nm in frame_text.keys():
                    if(frame_text[tb_nm].has_key(frame_n)):
                        rs.SetUserText(obj,tb_nm,frame_text[tb_nm][frame_n])
        #生成areas_dict
        for i in connectivity_area:
            points=[]
            area_n=-1
            for j in i:
                l=j.split("=")
                if (len(l)==2) and l[0]=="Area":
                    area_n=l[1]
                elif (len(l)==2) and re.match("Joint",l[0]):
                    if joints_dict.has_key(l[1]):
                        points.append(joints_dict[l[1]])
                    else:
                        continue
            if len(points)>=3 and (area_n!=-1):
                obj=rs.AddSrfPt(points)
                areas_dict[area_n]=obj
                for tb_nm in area_text.keys():
                    if(area_text[tb_nm].has_key(area_n)):
                        rs.SetUserText(obj,tb_nm,area_text[tb_nm][area_n])
    
    def export_mgt(self):
        "Export to a mgt text file"
        self.set_data()
        #prompt the user to specify a file name
        filter = "MGT File (*.mgt)|*.mgt|All files (*.*)|*.*||"
        filename = rs.SaveFileName("Save model As", filter)
        if not filename: return
    
        #open a new file
        file = open( filename, "w" )
        #document data中的表格
        for nm in self.m_mgt_docu_table_names:
            if(rs.GetDocumentData(nm)):
                file.write(nm)
                for entry in rs.GetDocumentData(nm):
                    file.write(entry)
                file.write("\n")

        # Nodes
        file.write("*NODE    ; Nodes\n")
        file.write("; iNO, X, Y, Z\n")
        for coord in self.m_joints_dict:
            joint=self.m_joints_dict[coord]
            file.write("     %s, %s, %s, %s\n"%(joint.m_number,joint.m_X,joint.m_Y,joint.m_Z))
        file.write("\n")
    
        # Elements
        file.write("*ELEMENT    ; Elements\n")
        file.write("; iEL, TYPE, iMAT, iPRO, iN1, iN2, ANGLE, iSUB, EXVAL, iOPT(EXVAL2) ; Frame  Element\n")
        file.write("; iEL, TYPE, iMAT, iPRO, iN1, iN2, ANGLE, iSUB, EXVAL, EXVAL2, bLMT ; Comp/Tens Truss\n")
        file.write("; iEL, TYPE, iMAT, iPRO, iN1, iN2, iN3, iN4, iSUB, iWID , LCAXIS    ; Planar Element\n")
        file.write("; iEL, TYPE, iMAT, iPRO, iN1, iN2, iN3, iN4, iN5, iN6, iN7, iN8     ; Solid  Element\n")
        for frame in self.m_frames:
            if rs.GetUserText(frame.m_GUID,"mgt_element") is not None:
                str=rs.GetUserText(frame.m_GUID,"mgt_element")
                ls=str.split(",")
                if len(ls)>=6:
                    ls[0]="  %s"%frame.m_number
                    ls[4]="   %s"%frame.m_start_joint.m_number
                    ls[5]="   %s"%frame.m_end_joint.m_number
                for i in range(len(ls)):
                    file.write(ls[i])
                    if(i!=len(ls)-1):
                        file.write(",")
            else:
                file.write("  %s, TRUSS ,    1,     1,   %s,   %s,     0,     0"%(frame.m_number,frame.m_start_joint.m_number,frame.m_end_joint.m_number))
            file.write("\n")
        file.write("\n")
    
        # EndData
        file.write("*ENDDATA\n")
    
        # Close File
        file.close()
        
    def import_mgt(self):
        #prompt the user for a file to import
        filter = "mgt file (*.mgt)|*.mgt|All Files (*.*)|*.*||"
        filename = rs.OpenFileName("Open MGT File", filter)
        if not filename: return
    
        #read each line from the file
        file = open(filename, "r")
        
        table=""
        joints_dict={} #以编号为key，以坐标元组为value
        trusses=[] #中间存储形式
        trusses_dict={}#以编号为key，以GUID为value
        for line in file:
            #去掉“;”后的注释部分
            ls=line.split(";")
            str=ls[0]
                
            #如果当前行是表名行，判断是哪个表名行
            if re.search("\\*",str) is not None:
                table=""
                if re.search("\\*NODE",str) is not None:
                    table="NODE"
                elif re.search("\\*ELEMENT",str) is not None:
                    table="ELEMENT"
                elif re.search("\\*ENDDATA",str) is not None:
                    break
                else:
                    table=line
                continue
                
            #如果当前行不是表名行
            #生成joints_dict和trusses
            #将document data直接存入
            if table=="":
                continue
            elif table=="NODE":
                str.replace(" ","")
                ls=str.split(",")
                if len(ls)>=4:
                    joints_dict[int(ls[0])]=(float(ls[1]),float(ls[2]),float(ls[3]))
            elif table=="ELEMENT":
                ls=str.split(",")
                if (len(ls)>=8) and (ls[1].replace(" ","")=="TRUSS"):
                    trusses.append(line)
            else:
                rs.SetDocumentData("mgt_docu_table_names",table," ")
                rs.SetDocumentData(table,line," ")

        file.close()
        #print(trusses)
        #print(joints_dict)
        #生成trusses_dict并在rhino中生成模型
        for text in trusses:
            ls=text.split(",")
            if len(ls)<6: continue
            if joints_dict.has_key(int(ls[4])) and joints_dict.has_key(int(ls[5])):
                obj=rs.AddLine(joints_dict[int(ls[4])],joints_dict[int(ls[5])])
                rs.SetUserText(obj,"mgt_element",text)
                trusses_dict[int(ls[0])]=obj
            
if( __name__ == "__main__" ):
    model=ModelInfo()
    
    #model.import_s2k()
    model.export_s2k()
    #model.import_mgt()
    #model.export_mgt()