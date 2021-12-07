#encoding=utf-8
import rhinoscriptsyntax as rs
from System.Windows.Forms import *
from System.Drawing import *
from System.Threading import ThreadStart, Thread
#import rhinoscript.selection
import rhinoscript.geometry
import re
import sys
import time

reload(sys)

sys.setdefaultencoding('utf-8')

class Joint:
    def __init__(self,number,X,Y,Z,GUID=""):
        self.m_number=number
        self.m_X=float(X)
        self.m_Y=float(Y)
        self.m_Z=float(Z)
        self.m_GlobalX=float(X)
        self.m_GlobalY=float(Y)
        self.m_GlobalZ=float(Z)
        self.m_GUID=GUID
        # 初始化时不从外部赋值的属性
        self.m_CoordSys="GLOBAL"
        self.m_CoordType="Cartesian"
        self.m_SpecialJt="No"
        self.m_Frames=[]
    
    #计算该节点出发的杆件间的最大最小夹角
    def cal_max_min_frame_angle(self):
        if self.m_Frames is None:
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
                if self.m_Frames[i].m_start_joint!=self:
                    vec1=-self.m_Frames[i].m_vector
                else:
                    vec1=self.m_Frames[i].m_vector
                for j in range(i+1,len(self.m_Frames)):
                    if self.m_Frames[j].m_start_joint!=self:
                        vec2=-self.m_Frames[j].m_vector
                    else:
                        vec2=self.m_Frames[j].m_vector
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
        self.m_layer="other"

class Area:
    def __init__(self,number,joint_count,joints,GUID):
        self.m_number=number
        self.m_joint_count=joint_count
        self.m_joints=joints
        self.m_GUID=GUID
        self.m_CentroidX=0
        self.m_CentroidY=0
        self.m_CentroidZ=0
        for j in self.m_joints:
            self.m_CentroidX+=j.m_X
            self.m_CentroidY+=j.m_Y
            self.m_CentroidZ+=j.m_Z
        self.m_CentroidX=self.m_CentroidX/float(joint_count)
        self.m_CentroidY=self.m_CentroidY/float(joint_count)
        self.m_CentroidZ=self.m_CentroidZ/float(joint_count)
        self.m_layer="other"

class DisplayGridInfoForm(Form):
    # build all of the controls in the constructor
    def __init__(self, max_angle, min_angle, max_length):
        offset = 10
        self.Text = "网格质量参数"
    
        max_angle_label = Label(Text="相邻杆件最大夹角 = %s"%max_angle, AutoSize=True)
        self.Controls.Add(max_angle_label)
        width = max_angle_label.Right
        pt = Point(max_angle_label.Left,max_angle_label.Bottom + offset)
    
        min_angle_label = Label(Text="相邻杆件最小夹角 = %s"%min_angle, AutoSize=True)
        min_angle_label.Location=pt
        self.Controls.Add(min_angle_label)
        if( min_angle_label.Right > width ):
            width = min_angle_label.Right
        pt = Point(min_angle_label.Left,min_angle_label.Bottom + offset)
    
        max_length_label = Label(Text="最长杆件长度 = %s"%max_length, AutoSize=True)
        max_length_label.Location=pt
        self.Controls.Add(max_length_label)
        if( max_length_label.Right > width ):
            width = max_length_label.Right
        pt = Point(max_length_label.Left,max_length_label.Bottom + offset)
    
        buttonApply = Button(Text="显示节点", DialogResult=DialogResult.OK)
        buttonApply.Location = pt
        self.Controls.Add(buttonApply)
        pt.X = buttonApply.Right + offset
        buttonCancel = Button(Text="取消", DialogResult=DialogResult.Cancel)
        buttonCancel.Location = pt
        self.Controls.Add(buttonCancel)
        if( buttonCancel.Right > width ):
            width = buttonCancel.Right
        self.ClientSize = Size(width, buttonCancel.Bottom)
        self.AcceptButton = buttonApply
        self.CancelButton = buttonCancel

        
class ProgressBarForm(Form):
    def __init__(self,label,max_step):
        self.offset = 10
        self.Text = "Progress Bar"

        self.task_label=Label(Text=label, AutoSize=True) #label为任务名称：可定义为“导出mgt文件”“导出s2k文件”
        self.Controls.Add(self.task_label)
        width = self.task_label.Right
        pt = Point(self.task_label.Left,self.task_label.Bottom + self.offset)
        
        self.pbar=ProgressBar(Maximum=max_step,Visible=True,Step=1,Value=0)
        self.pbar.Location=pt
        self.pbar.Size=Size(800,20)
        self.Controls.Add(self.pbar)
        
        if( self.pbar.Right > width ):
            width = self.pbar.Right
            
        self.ClientSize = Size(width, self.pbar.Bottom)
        self.invoker=MethodInvoker(self.update)
    
    def update(self):
        self.pbar.PerformStep()
        
        
class ModelInfo:
    def __init__(self):
        #构件数据：object data
        self.m_frames=[]
        self.m_areas=[]
        self.m_joints_dict={}#以坐标为kee，Joint对象为value
       
        #以下属性表名从文本文件中整行读取，不进行去空格和去换行符处理
        #导出时也直接照读取进来的原样打印不用加换行符
        self.m_s2k_joint_table_names=[]
        self.m_s2k_frame_table_names=[]#s2k中每根杆件对应的属性表名
        self.m_s2k_area_table_names=[]#s2k中每根杆件对应的属性表名
        self.m_mgt_obj_table_names=[]#mgt中每根杆件对应的属性表名
        self.m_s2k_docu_table_names=[]#s2k中其他属性表名
        self.m_mgt_docu_table_names=[]#mgt中其他属性表名
        
        self.form=None
        
    def pbar_thread(self):
        self.form.Show()
        Application.Run(self.form)
        
    def set_data_mgt(self):
        objectIds = rs.GetObjects("Select")
        #print(len(objectIds))
        if objectIds is None: return
        j_n=0
        e_n=0
        max_j_n=0
        max_e_n=0
        if (rs.GetDocumentData("mgt_Node_original_number","max_number")):
            max_j_n=int(rs.GetDocumentData("mgt_Node_original_number","max_number"))
        if (rs.GetDocumentData("mgt_Element_original_number","max_number")):
            max_e_n=int(rs.GetDocumentData("mgt_Element_original_number","max_number"))
            
        for objectId in objectIds:
            # 将端点存入字典,键类型为代表坐标的三维元组，值类型为joint
            #如果类型为4（Curve）
            if rs.ObjectType(objectId)==4:
                start_x,start_y,start_z=rs.CurveStartPoint(objectId)
                start_coord=(start_x,start_y,start_z)
                if self.m_joints_dict.has_key(start_coord):
                    start_joint=self.m_joints_dict[start_coord]
                else:
                    #if(rs.GetDocumentData("mgt_Node_original_number","(%.11f,%.11f,%.11f)"%(start_x,start_y,start_z))):
                        #j_n=rs.GetDocumentData("mgt_Node_original_number","(%.11f,%.11f,%.11f)"%(start_x,start_y,start_z))
                    if rs.GetUserText(objectId,"start_point_guid"):
                        j_n=int(rs.GetUserText(rs.GetUserText(objectId,"start_point_guid"),"mgt_original_number"))
                    else:
                        max_j_n+=1
                        j_n=max_j_n
                    start_joint=Joint(j_n,start_x,start_y,start_z)
                    self.m_joints_dict[start_coord]=start_joint
                end_x,end_y,end_z=rs.CurveEndPoint(objectId)
                end_coord=(end_x,end_y,end_z)
                if self.m_joints_dict.has_key(end_coord):
                    end_joint=self.m_joints_dict[end_coord]
                else:
                    #如果是从mgt中导入的，就使用原来的编号
                    #if(rs.GetDocumentData("mgt_Node_original_number","(%.11f,%.11f,%.11f)"%(end_x,end_y,end_z))):
                        #j_n=rs.GetDocumentData("mgt_Node_original_number","(%.11f,%.11f,%.11f)"%(end_x,end_y,end_z))
                    if rs.GetUserText(objectId,"end_point_guid"):
                        j_n=int(rs.GetUserText(rs.GetUserText(objectId,"end_point_guid"),"mgt_original_number"))
                    #否则使用新编号，最大编号加一
                    else:
                        max_j_n+=1
                        j_n=max_j_n
                    end_joint=Joint(j_n,end_x,end_y,end_z)
                    self.m_joints_dict[end_coord]=end_joint
                # 将杆件存入列表
                if(rs.GetUserText(objectId,"mgt_original_number")):
                    e_n=int(rs.GetUserText(objectId,"mgt_original_number"))
                else:
                    max_e_n+=1
                    e_n=max_e_n
                frame=Frame(e_n,start_joint,end_joint,objectId)
                self.m_frames.append(frame)
                self.m_joints_dict[start_coord].m_Frames.append(frame)
                self.m_joints_dict[end_coord].m_Frames.append(frame)
                #判断是否有s2k section属性，如果没有，将layer name作为section name存入
                #if (rs.GetUserText(objectId,"s2k_section_layer_name") is None) or (rs.GetUserText(objectId,"s2k_section_layer_name")!=rs.ObjectLayer(objectId)):
                    #sec_nm=rs.ObjectLayer(objectId)
                    #rs.SetUserText(objectId,"TABLE:  \"FRAME SECTION ASSIGNMENTS\"\n","   Frame=%s   AutoSelect=N.A.   AnalSect=%s   MatProp=Default\n"%(f_n,sec_nm))
                    #rs.SetDocumentData("TABLE:  \"FRAME SECTION PROPERTIES 01 - GENERAL\"\n","   SectionName=%s   Shape=\"I/Wide Flange\"   Color=Red\n"%sec_nm,"   SectionName=%s   Shape=\"I/Wide Flange\"   Color=Red\n"%sec_nm)
                
            #如果类型为8（Surface）
            elif rs.ObjectType(objectId)==8:
                area_joints=[]
                points=rs.SurfacePoints(objectId)
                #if(len(points)>4)|(len(points)<3):
                    #continue
                #elif (len(points)==3):
                    #for i in [0,1:
                        #x,y,z=pt
                        #coord=(x,y,z)
                        #if self.m_joints_dict.has_key(coord):
                            #area_joints.append(self.m_joints_dict[coord])
                        #else:
                            #if(rs.GetDocumentData("s2k_Joint_original_number","(%.11f,%.11f,%.11f)"%(x,y,z))):
                                #j_n=rs.GetDocumentData("s2k_Joint_original_number","(%.11f,%.11f,%.11f)"%(x,y,z))
                            #else:
                                #max_j_n+=1
                                #j_n=max_j_n
                            #tmp_j=Joint(j_n,x,y,z)
                            #self.m_joints_dict[coord]=tmp_j
                            #area_joints.append(tmp_j)
                            
                #如果面对象中存有节点对应点对象的GUID列表，则读取原有节点编号
                if rs.GetUserText(objectId,"point_guids"):
                    pt_guid_list=rs.GetUserText(objectId,"point_guids").split(",")
                    for i in range(len(pt_guid_list)):
                        x,y,z=points[i]
                        coord=(x,y,z)
                        if self.m_joints_dict.has_key(coord):
                            if self.m_joints_dict[coord] not in area_joints:
                                area_joints.append(self.m_joints_dict[coord])
                        else:
                            if rs.GetUserText(pt_guid_list[i],"mgt_original_number"):
                                j_n=int(rs.GetUserText(pt_guid_list[i],"mgt_original_number"))
                            else:
                                max_j_n+=1
                                j_n=max_j_n
                            tmp_j=Joint(j_n,x,y,z)
                            self.m_joints_dict[coord]=tmp_j
                            if tmp_j not in area_joints:
                                area_joints.append(tmp_j)
                #如果面对象中没有存节点对应点对象GUID，则需要新分配节点编号
                else:
                    for i in range(len(points)):
                        x,y,z=points[i]
                        coord=(x,y,z)
                        if self.m_joints_dict.has_key(coord):
                            if self.m_joints_dict[coord] not in area_joints:
                                area_joints.append(self.m_joints_dict[coord])
                        else:
                            max_j_n+=1
                            j_n=max_j_n
                            tmp_j=Joint(j_n,x,y,z)
                            self.m_joints_dict[coord]=tmp_j
                            if tmp_j not in area_joints:
                                area_joints.append(tmp_j)
                                
                if(rs.GetUserText(objectId,"mgt_original_number")):
                    e_n=int(rs.GetUserText(objectId,"mgt_original_number"))
                else:
                    max_e_n+=1
                    e_n=max_e_n
                j_c=len(area_joints)
                area=Area(e_n,j_c,area_joints,objectId)
                self.m_areas.append(area)
                #判断是否有section属性，如果没有，将layer name作为section name存入
                #if rs.GetUserText(objectId,"TABLE:  \"AREA SECTION ASSIGNMENTS\"\n") is None:
                    #sec_nm=rs.ObjectLayer(objectId)
                    #rs.SetUserText(objectId,"TABLE:  \"AREA SECTION ASSIGNMENTS\"\n","   Area=%s   Section=%s   MatProp=Default\n"%(a_n,sec_nm))
                    #rs.SetDocumentData("TABLE:  \"AREA SECTION PROPERTIES\"\n","   SectionName=%s   AreaType=Shell   Color=Green\n"%sec_nm,"   SectionName=%s   AreaType=Shell   Color=Green\n"%sec_nm)
            
        #读取documentdata中存下来的表名
        #if(rs.GetDocumentData("s2k_frame_table_names")):
            #rs.SetDocumentData("s2k_frame_table_names","TABLE:  \"FRAME SECTION ASSIGNMENTS\"\n","TABLE:  \"FRAME SECTION ASSIGNMENTS\"\n")
            #self.m_s2k_frame_table_names=rs.GetDocumentData("s2k_frame_table_names")
        #if(rs.GetDocumentData("s2k_area_table_names")):
            #rs.SetDocumentData("s2k_area_table_names","TABLE:  \"AREA SECTION ASSIGNMENTS\"\n","TABLE:  \"AREA SECTION ASSIGNMENTS\"\n")
            #self.m_s2k_area_table_names=rs.GetDocumentData("s2k_area_table_names")
        if(rs.GetDocumentData("mgt_obj_table_names")):
            self.m_mgt_obj_table_names=rs.GetDocumentData("mgt_obj_table_names")
        #if(rs.GetDocumentData("s2k_docu_table_names")):
            #rs.SetDocumentData("s2k_docu_table_names","TABLE:  \"AREA SECTION PROPERTIES\"\n","TABLE:  \"AREA SECTION PROPERTIES\"\n")
            #rs.SetDocumentData("s2k_docu_table_names","TABLE:  \"FRAME SECTION PROPERTIES 01 - GENERAL\"\n","TABLE:  \"FRAME SECTION PROPERTIES 01 - GENERAL\"\n")
            #self.m_s2k_docu_table_names=rs.GetDocumentData("s2k_docu_table_names")
        if(rs.GetDocumentData("mgt_docu_table_names")):
            self.m_mgt_docu_table_names=rs.GetDocumentData("mgt_docu_table_names")
        #print("Frame_Count=%s"%len(self.m_frames))
        return
     
    def set_data_s2k(self):
        objectIds = rs.GetObjects("Select")
        #print(len(objectIds))
        if objectIds is None: return
        j_n=0
        f_n=0
        a_n=0
        max_j_n=0
        max_f_n=0
        max_a_n=0
        if (rs.GetDocumentData("s2k_Joint_original_number","max_number")):
            max_j_n=int(rs.GetDocumentData("s2k_Joint_original_number","max_number"))
        if (rs.GetDocumentData("s2k_Frame_original_number","max_number")):
            max_f_n=int(rs.GetDocumentData("s2k_Frame_original_number","max_number"))
        if (rs.GetDocumentData("s2k_Area_original_number","max_number")):
            max_a_n=int(rs.GetDocumentData("s2k_Area_original_number","max_number"))
            
        for objectId in objectIds:
            # 将端点存入字典,键类型为代表坐标的三维元组，值类型为joint
            #如果类型为4（Curve）
            if rs.ObjectType(objectId)==4:
                start_x,start_y,start_z=rs.CurveStartPoint(objectId)
                start_coord=(start_x,start_y,start_z)
                if self.m_joints_dict.has_key(start_coord):
                    start_joint=self.m_joints_dict[start_coord]
                else:
                    joint_guid=""
                    if rs.GetUserText(objectId,"start_point_guid"):
                        if rs.GetUserText(rs.GetUserText(objectId,"start_point_guid"),"s2k_original_number"):
                            j_n=int(rs.GetUserText(rs.GetUserText(objectId,"start_point_guid"),"s2k_original_number"))
                            joint_guid=rs.GetUserText(objectId,"start_point_guid")
                        else:
                            max_j_n+=1
                            j_n=max_j_n
                    #if(rs.GetDocumentData("s2k_Joint_original_number","(%.11f,%.11f,%.11f)"%(start_x,start_y,start_z))):
                        #j_n=rs.GetDocumentData("s2k_Joint_original_number","(%.11f,%.11f,%.11f)"%(start_x,start_y,start_z))
                    else:
                        max_j_n+=1
                        j_n=max_j_n
                    start_joint=Joint(j_n,start_x,start_y,start_z,GUID=joint_guid)
                    self.m_joints_dict[start_coord]=start_joint
                    
                end_x,end_y,end_z=rs.CurveEndPoint(objectId)
                end_coord=(end_x,end_y,end_z)
                if self.m_joints_dict.has_key(end_coord):
                    end_joint=self.m_joints_dict[end_coord]
                else:
                    #如果是从s2k中导入的，就使用原来的编号
                    joint_guid=""
                    if rs.GetUserText(objectId,"end_point_guid"):
                        if rs.GetUserText(rs.GetUserText(objectId,"end_point_guid"),"s2k_original_number"):
                            j_n=int(rs.GetUserText(rs.GetUserText(objectId,"end_point_guid"),"s2k_original_number"))
                            joint_guid=rs.GetUserText(objectId,"end_point_guid")
                        else:
                            max_j_n+=1
                            j_n=max_j_n
                    #if(rs.GetDocumentData("s2k_Joint_original_number","(%.11f,%.11f,%.11f)"%(end_x,end_y,end_z))):
                        #j_n=rs.GetDocumentData("s2k_Joint_original_number","(%.11f,%.11f,%.11f)"%(end_x,end_y,end_z))
                       
                    #否则使用新编号，最大编号加一
                    else:
                        max_j_n+=1
                        j_n=max_j_n
                    end_joint=Joint(j_n,end_x,end_y,end_z,GUID=joint_guid)
                    self.m_joints_dict[end_coord]=end_joint
                # 将杆件存入列表
                if(rs.GetUserText(objectId,"s2k_original_number")):
                    f_n=rs.GetUserText(objectId,"s2k_original_number")
                elif(rs.GetUserText(objectId,"mgt_original_number")):
                    f_n=rs.GetUserText(objectId,"mgt_original_number")
                else:
                    max_f_n+=1
                    f_n=max_f_n
                frame=Frame(f_n,start_joint,end_joint,objectId)
                self.m_frames.append(frame)
                self.m_joints_dict[start_coord].m_Frames.append(frame)
                self.m_joints_dict[end_coord].m_Frames.append(frame)
                #判断是否有s2k section属性，如果没有，将layer name作为section name存入
                if (rs.GetUserText(objectId,"s2k_section_layer_name") is None) or (rs.GetUserText(objectId,"s2k_section_layer_name")!=rs.ObjectLayer(objectId)):
                    sec_nm=rs.ObjectLayer(objectId)
                    rs.SetUserText(objectId,"TABLE:  \"FRAME SECTION ASSIGNMENTS\"\n","   Frame=%s   AutoSelect=N.A.   AnalSect=%s   MatProp=Default\n"%(f_n,sec_nm))
                    rs.SetDocumentData("TABLE:  \"FRAME SECTION PROPERTIES 01 - GENERAL\"\n","   SectionName=%s   Shape=\"I/Wide Flange\"   Color=Red\n"%sec_nm,"   SectionName=%s   Shape=\"I/Wide Flange\"   Color=Red\n"%sec_nm)
                
            #如果类型为8（Surface）
            elif rs.ObjectType(objectId)==8:
                area_joints=[]
                points=rs.SurfacePoints(objectId)
                #if(len(points)>4)|(len(points)<3):
                    #continue
                #elif (len(points)==3):
                    #for pt in points:
                        #x,y,z=pt
                        #coord=(x,y,z)
                        #if self.m_joints_dict.has_key(coord):
                            #area_joints.append(self.m_joints_dict[coord])
                        #else:
                            #if(rs.GetDocumentData("s2k_Joint_original_number","(%.11f,%.11f,%.11f)"%(x,y,z))):
                                #j_n=rs.GetDocumentData("s2k_Joint_original_number","(%.11f,%.11f,%.11f)"%(x,y,z))
                            #else:
                                #max_j_n+=1
                                #j_n=max_j_n
                            #tmp_j=Joint(j_n,x,y,z)
                            #self.m_joints_dict[coord]=tmp_j
                            #area_joints.append(tmp_j)
                            
                if rs.GetUserText(objectId,"point_guids"):
                    pt_guid_list=rs.GetUserText(objectId,"point_guids").split(",")
                    tmp_index=[0,2,1]
                    if len(pt_guid_list)==4:
                        tmp_index=[0,2,3,1]
                        
                    for i in tmp_index:
                        x,y,z=points[i]
                        coord=(x,y,z)
                        if self.m_joints_dict.has_key(coord):
                            if self.m_joints_dict[coord] not in area_joints:
                                area_joints.append(self.m_joints_dict[coord])
                        else:
                            joint_guid=""
                            if rs.GetUserText(pt_guid_list[i],"s2k_original_number"):
                                j_n=int(rs.GetUserText(pt_guid_list[i],"s2k_original_number"))
                                joint_guid=pt_guid_list[i]
                            else:
                                max_j_n+=1
                                j_n=max_j_n
                            tmp_j=Joint(j_n,x,y,z,GUID=joint_guid)
                            self.m_joints_dict[coord]=tmp_j
                            if tmp_j not in area_joints:
                                area_joints.append(tmp_j)
                #如果面对象中没有存节点对应点对象GUID，则需要新分配节点编号
                else:
                    for i in range(len(points)):
                        x,y,z=points[i]
                        coord=(x,y,z)
                        if self.m_joints_dict.has_key(coord):
                            if self.m_joints_dict[coord] not in area_joints:
                                area_joints.append(self.m_joints_dict[coord])
                        else:
                            max_j_n+=1
                            j_n=max_j_n
                            tmp_j=Joint(j_n,x,y,z)
                            self.m_joints_dict[coord]=tmp_j
                            if tmp_j not in area_joints:
                                area_joints.append(tmp_j)
                #else:
                    #for i in [0,1,3,2]:
                        #x,y,z=points[i]
                        #coord=(x,y,z)
                        #if self.m_joints_dict.has_key(coord):
                            #area_joints.append(self.m_joints_dict[coord])
                        #else:
                            #if(rs.GetDocumentData("s2k_Joint_original_number","(%.11f,%.11f,%.11f)"%(x,y,z))):
                                #j_n=rs.GetDocumentData("s2k_Joint_original_number","(%.11f,%.11f,%.11f)"%(x,y,z))
                            #else:
                                #max_j_n+=1
                                #j_n=max_j_n
                            #tmp_j=Joint(j_n,x,y,z)
                            #self.m_joints_dict[coord]=tmp_j
                            #area_joints.append(tmp_j)
                            

                if(rs.GetUserText(objectId,"s2k_original_number")):
                    a_n=rs.GetUserText(objectId,"s2k_original_number")
                else:
                    max_a_n+=1
                    a_n=max_a_n
                j_c=len(area_joints)
                area=Area(a_n,j_c,area_joints,objectId)
                self.m_areas.append(area)
                #判断是否有section属性，如果没有，将layer name作为section name存入
                if rs.GetUserText(objectId,"TABLE:  \"AREA SECTION ASSIGNMENTS\"\n") is None:
                    sec_nm=rs.ObjectLayer(objectId)
                    rs.SetUserText(objectId,"TABLE:  \"AREA SECTION ASSIGNMENTS\"\n","   Area=%s   Section=%s   MatProp=Default\n"%(a_n,sec_nm))
                    rs.SetDocumentData("TABLE:  \"AREA SECTION PROPERTIES\"\n","   SectionName=%s   AreaType=Shell   Color=Green\n"%sec_nm,"   SectionName=%s   AreaType=Shell   Color=Green\n"%sec_nm)
            
        #读取documentdata中存下来的表名
        if(rs.GetDocumentData("s2k_joint_table_names")):
            self.m_s2k_joint_table_names=rs.GetDocumentData("s2k_joint_table_names")
        if(rs.GetDocumentData("s2k_frame_table_names")):
            rs.SetDocumentData("s2k_frame_table_names","TABLE:  \"FRAME SECTION ASSIGNMENTS\"\n","TABLE:  \"FRAME SECTION ASSIGNMENTS\"\n")
            self.m_s2k_frame_table_names=rs.GetDocumentData("s2k_frame_table_names")
        if(rs.GetDocumentData("s2k_area_table_names")):
            rs.SetDocumentData("s2k_area_table_names","TABLE:  \"AREA SECTION ASSIGNMENTS\"\n","TABLE:  \"AREA SECTION ASSIGNMENTS\"\n")
            self.m_s2k_area_table_names=rs.GetDocumentData("s2k_area_table_names")
        #if(rs.GetDocumentData("mgt_obj_table_names")):
            #self.m_mgt_obj_table_names=rs.GetDocumentData("mgt_obj_table_names")
        if(rs.GetDocumentData("s2k_docu_table_names")):
            rs.SetDocumentData("s2k_docu_table_names","TABLE:  \"AREA SECTION PROPERTIES\"\n","TABLE:  \"AREA SECTION PROPERTIES\"\n")
            rs.SetDocumentData("s2k_docu_table_names","TABLE:  \"FRAME SECTION PROPERTIES 01 - GENERAL\"\n","TABLE:  \"FRAME SECTION PROPERTIES 01 - GENERAL\"\n")
            self.m_s2k_docu_table_names=rs.GetDocumentData("s2k_docu_table_names")
        #if(rs.GetDocumentData("mgt_docu_table_names")):
            #self.m_mgt_docu_table_names=rs.GetDocumentData("mgt_docu_table_names")
        #print("Frame_Count=%s"%len(self.m_frames))
        return
    
    def export_s2k(self):
        #set data
        self.set_data_s2k()
        #prompt the user to specify a file name
        filter = "s2k File (*.s2k)|*.s2k|All files (*.*)|*.*||"
        filename = rs.SaveFileName("Save model Points As", filter)
        if not filename: return
        
        #进度条窗口类
        self.form=ProgressBarForm("导出s2k文件中",8)
        th=Thread(ThreadStart(self.pbar_thread))
        th.IsBackground=True
        th.Start()
    
        #open a new file
        file = open( filename, "w" )
        # file path
        file.write("File %s was saved on m/d/yy at h:mm:ss\n"%filename)
        file.write("\n")
        
        #table program control
        file.write("TABLE:  \"PROGRAM CONTROL\"\n")
        if rs.GetDocumentData("TABLE:  \"PROGRAM CONTROL\"\n"):
            for entry in rs.GetDocumentData("TABLE:  \"PROGRAM CONTROL\"\n"):
                file.write(entry)
        else:
            file.write("   ProgramName=SAP2000   Version=21.0.2   ProgLevel=Ultimate   LicenseNum=0   LicenseOS=Yes   LicenseSC=Yes   LicenseHT=No   CurrUnits=\"KN, m, C\"   SteelCode=\"Chinese 2010\"   ConcCode=\"Chinese 2010\"   AlumCode=\"AA-ASD 2000\"   ColdCode=AISI-ASD96   RegenHinge=Yes")
        file.write("\n")
        self.form.Invoke(self.form.invoker)
        
        for nm in self.m_s2k_docu_table_names:
            if(rs.GetDocumentData(nm)):
                file.write(nm)
                for entry in rs.GetDocumentData(nm):
                    file.write(entry)
                file.write("\n")
        
        self.form.Invoke(self.form.invoker)
        
        
        #以下3张为特殊表格
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
        self.form.Invoke(self.form.invoker)
        #table CONNECTIVITY - FRAME
        file.write("TABLE:  \"CONNECTIVITY - FRAME\"\n")
        print("joints_count=%s"%len(self.m_joints_dict))
        print("frames_count=%s"%len(self.m_frames))
        for frame in self.m_frames:
            file.write("   Frame=%s"%frame.m_number)
            file.write("   JointI=%s   JointJ=%s"%(frame.m_start_joint.m_number,frame.m_end_joint.m_number))
            file.write("   IsCurved=%s"%frame.m_IsCurved)
            file.write("   Length=%s"%frame.m_Length)
            file.write("   CentroidX=%s   CentroidY=%s   CentroidZ=%s"%(frame.m_CentroidX,frame.m_CentroidY,frame.m_CentroidZ))
            file.write("   GUID=%s"%frame.m_GUID)
            file.write("\n")
        file.write("\n")
        self.form.Invoke(self.form.invoker)
        #table CONNECTIVITY - AREA
        file.write("TABLE:  \"CONNECTIVITY - AREA\"\n")
        for area in self.m_areas:
            file.write("   Area=%s"%(area.m_number))
            file.write("   NumJoints=%s"%area.m_joint_count)
            for i in range(area.m_joint_count):
                file.write("   Joint%s=%s"%(i+1,area.m_joints[i].m_number))
            file.write("   Perimeter=12.0551750201988   AreaArea=8.11035004039763   Volume=0.811035004039762")
            file.write("   CentroidX=%s   CentroidY=%s   CentroidZ=%s   GUID=%s"%(area.m_CentroidX,area.m_CentroidY,area.m_CentroidZ,area.m_GUID))
            file.write("\n")
        file.write("\n")
        self.form.Invoke(self.form.invoker)
        #form.pbar.PerformStep()
        
        for nm in self.m_s2k_joint_table_names:
            file.write(nm)
            #print(nm)
            for jt_coord in self.m_joints_dict:
                jt=self.m_joints_dict[jt_coord]
                if jt.m_GUID=="":
                    continue
                if rs.GetUserText(jt.m_GUID,nm) is None:
                    continue
                #print(rs.GetUserText(fm.m_GUID,nm))
                str_tmp=rs.GetUserText(jt.m_GUID,nm)
                l=str_tmp.split()
                res=re.search("Joint=",str_tmp)
                if (len(l)<1) | (res is None):
                    continue
                a=str(jt.m_number)
                file.write(str_tmp[:res.end()]+a+str_tmp[res.end()+len(l[0])-6:])
            file.write("\n")
        self.form.Invoke(self.form.invoker)
            
        for nm in self.m_s2k_frame_table_names:
            file.write(nm)
            #print(nm)
            for fm in self.m_frames:
                if rs.GetUserText(fm.m_GUID,nm) is None:
                    continue
                #print(rs.GetUserText(fm.m_GUID,nm))
                str_tmp=rs.GetUserText(fm.m_GUID,nm)
                l=str_tmp.split()
                res=re.search("Frame=",str_tmp)
                if (len(l)<1) | (res is None):
                    continue
                a=str(fm.m_number)
                file.write(str_tmp[:res.end()]+a+str_tmp[res.end()+len(l[0])-6:])
            file.write("\n")
        self.form.Invoke(self.form.invoker)
        
        for nm in self.m_s2k_area_table_names:
            file.write(nm)
            #print(nm)
            for ar in self.m_areas:
                if rs.GetUserText(ar.m_GUID,nm) is None:
                    continue
                str_tmp=rs.GetUserText(ar.m_GUID,nm)
                l=str_tmp.split()
                res=re.search("Area=",str_tmp)
                if (len(l)<1) | (res is None):
                    continue
                a=str(ar.m_number)
                file.write(str_tmp[:res.end()]+a+str_tmp[res.end()+len(l[0])-5:])
            file.write("\n")
        self.form.Invoke(self.form.invoker)
        file.write("END TABLE DATA")
        #close the file
        file.close()
        self.form.Close()
        #print(self.m_s2k_frame_table_names)
        #print("\n")
        #print(self.m_s2k_area_table_names)
        #print("\n")
        #print(self.m_s2k_docu_table_names)
        print("导出s2k文件完成！")
    
    def import_s2k(self):
        #prompt the user for a file to import
        filter = "s2k file (*.s2k)|*.s2k|All Files (*.*)|*.*||"
        filename = rs.OpenFileName("Open Point File", filter)
        if not filename: return
        
        self.form = ProgressBarForm("导入s2k文件中", 4)
        t = Thread(ThreadStart(self.pbar_thread))
        t.IsBackground = True
        t.Start()
        
        #read each line from the file
        file = open(filename, "r")
        table=""
        joint_coordinates=[]
        connectivity_frame=[]
        connectivity_area=[]
        joint_text={}
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
                unfinished_line=unfinished_line.replace("_","")
                unfinished_line=unfinished_line.replace("\n","")
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
                if (len(ls)>=4) and (ls[1]=="\"CONNECTIVITY") and (ls[2]=="-") and (ls[3]=="FRAME\""):
                    table="TABLE:  \"CONNECTIVITY - FRAME\"\n"
                    continue
                #table CONNECTIVITY - AREA
                if (len(ls)>=4) and (ls[1]=="\"CONNECTIVITY") and (ls[2]=="-") and (ls[3]=="FRAME\""):
                    table="TABLE:  \"CONNECTIVITY - AREA\"\n"
                    continue
                #section表，用于分配图层
                #TABLE:  "FRAME SECTION ASSIGNMENTS"
                if (len(ls)>=4) and (ls[1]=="\"FRAME") and (ls[2]=="SECTION") and (ls[3]=="ASSIGNMENTS\""):
                    table="TABLE:  \"FRAME SECTION ASSIGNMENTS\"\n"
                    continue
                if (len(ls)>=4) and (ls[1]=="\"AREA") and (ls[2]=="SECTION") and (ls[3]=="ASSIGNMENTS\""):
                    table="TABLE:  \"AREA SECTION ASSIGNMENTS\"\n"
                    continue
                #group表
                if (len(ls)>=2) and (ls[1]=="\"GROUPS") and (ls[-1]=="ASSIGNMENTS\""):
                    table="TABLE:  \"GROUPS 2 - ASSIGNMENTS\"\n"
                    continue
                    
                #其他表格
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
                    #是跟随每个Joint的text，存入一个中间结构joint_text
                    if len(tmp)>=2 and tmp[0]=="Joint":
                        if not joint_text.has_key(table):
                            tmp_dic={}
                            joint_text[table]=tmp_dic
                            rs.SetDocumentData("s2k_joint_table_names",str(table),str(table))
                        joint_text[table][tmp[1]]=whole_line
                    #是跟随每个Frame的text，存入一个中间结构frame_text
                    elif len(tmp)>=2 and tmp[0]=="Frame":
                        if not frame_text.has_key(table):
                            tmp_dic={}
                            frame_text[table]=tmp_dic
                            rs.SetDocumentData("s2k_frame_table_names",str(table),str(table))
                        frame_text[table][tmp[1]]=whole_line
                    #是跟随每个Area的text，存入一个中间结构area_text
                    elif len(tmp)>=2 and tmp[0]=="Area":
                        if not area_text.has_key(table):
                            tmp_dic={}
                            area_text[table]=tmp_dic
                            rs.SetDocumentData("s2k_area_table_names",str(table),str(table))
                        area_text[table][tmp[1]]=whole_line
                            
                    #是document data 直接存
                    else:
                        if (table!="TABLE:  \"PROGRAM CONTROL\"\n"):
                            rs.SetDocumentData("s2k_docu_table_names",str(table),str(table))
                        rs.SetDocumentData(table,whole_line,ls[0])
        file.close()
        self.form.Invoke(self.form.invoker)
    
        joints_dict={} #以编号为key，以GUID为value
        frames_dict={} #以编号为key，以GUID为value
        areas_dict={}  #以编号为key，以GUID为value
        #生成joints_dict
        max_number=0       
        for i in joint_coordinates:

            number=-1
            X=0.0
            Y=0.0
            Z=0.0
            for j in i:
                l=j.split("=")
                if (len(l)==2) and l[0]=="Joint":
                    number=l[1]
                    max_number=max(int(number),max_number)
                elif (len(l)==2) and l[0]=="XorR":
                    X=float(l[1])
                elif (len(l)==2) and l[0]=="Y":
                    Y=float(l[1])
                elif (len(l)==2) and l[0]=="Z":
                    Z=float(l[1])
            if number!=-1:
                obj_pt=rs.AddPoint((X,Y,Z))
                joints_dict[number]=obj_pt
                rs.SetUserText(obj_pt,"s2k_original_number",str(number))
                #rs.SetDocumentData("s2k_Joint_original_number","(%.11f,%.11f,%.11f)"%(X,Y,Z),number)
                for tb_nm in joint_text.keys():
                    if(joint_text[tb_nm].has_key(number)):
                        rs.SetUserText(obj_pt,tb_nm,joint_text[tb_nm][number])
        rs.SetDocumentData("s2k_Joint_original_number","max_number",str(max_number))
        #生成frames_dict
        #print(len(connectivity_frame))
        max_f_n=0
        self.form.Invoke(self.form.invoker)
        
        for i in connectivity_frame:
            jointi=-1
            jointj=-1
            frame_n=-1
            for j in i:
                l=j.split("=")
                if (len(l)==2) and l[0]=="Frame":
                    frame_n=l[1]
                    if(int(frame_n)>max_f_n):
                        max_f_n=int(frame_n)
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
                rs.SetUserText(obj,"s2k_original_number",frame_n)
                rs.SetUserText(obj,"start_point_guid",jointi)
                rs.SetUserText(obj,"end_point_guid",jointj)
                frames_dict[frame_n]=obj
                for tb_nm in frame_text.keys():
                    if(frame_text[tb_nm].has_key(frame_n)):
                        rs.SetUserText(obj,tb_nm,frame_text[tb_nm][frame_n])
                        #使用section进行图层分类
                        if tb_nm=="TABLE:  \"FRAME SECTION ASSIGNMENTS\"\n":
                            for i in frame_text[tb_nm][frame_n].split():
                                if i.split("=")[0]=="AnalSect":
                                    rs.AddLayer(i.split("=")[1])
                                    rs.ObjectLayer(obj,i.split("=")[1])
                                    rs.SetUserText(obj,"s2k_section_layer_name",i.split("=")[1])
                                    break
        rs.SetDocumentData("s2k_Frame_original_number","max_number",str(max_f_n))
        self.form.Invoke(self.form.invoker)
        #生成areas_dict
        max_a_n=0
        for i in connectivity_area:
            points=[]
            area_n=-1
            for j in i:
                l=j.split("=")
                if (len(l)==2) and l[0]=="Area":
                    area_n=l[1]
                    if(int(area_n)>max_a_n):
                        max_a_n=int(area_n)
                elif (len(l)==2) and re.match("Joint",l[0]):
                    if joints_dict.has_key(l[1]):
                        points.append(joints_dict[l[1]])
                    else:
                        continue
            if len(points)>=3 and (area_n!=-1):
                obj=rs.AddSrfPt(points)
                #将点对象guid存成一个字符串，按0，-1，1，2顺序存储
                tmp_str=str(points[0])+","+str(points[-1])+","+str(points[1])
                if len(points)>=4:
                    tmp_str=tmp_str+","+str(points[2])
                rs.SetUserText(obj,"point_guids",tmp_str)
                rs.SetUserText(obj,"s2k_original_number",area_n)
                areas_dict[area_n]=obj
                for tb_nm in area_text.keys():
                    if(area_text[tb_nm].has_key(area_n)):
                        rs.SetUserText(obj,tb_nm,area_text[tb_nm][area_n])
                        #使用section进行图层分类
                        if tb_nm=="TABLE:  \"AREA SECTION ASSIGNMENTS\"\n":
                            for i in area_text[tb_nm][area_n].split():
                                if i.split("=")[0]=="Section":
                                    rs.AddLayer(i.split("=")[1])
                                    rs.ObjectLayer(obj,i.split("=")[1])
                                    rs.SetUserText(obj,"s2k_section_layer_name",i.split("=")[1])
                                    break
        rs.SetDocumentData("s2k_Area_original_number","max_number",str(max_a_n))
        self.form.Invoke(self.form.invoker)
        self.form.Close()
        print("从s2k文件中导入完成！")
        
    def export_mgt(self):
        "Export to a mgt text file"
        self.set_data_mgt()
        #prompt the user to specify a file name
        filter = "MGT File (*.mgt)|*.mgt|All files (*.*)|*.*||"
        filename = rs.SaveFileName("Save model As", filter)
        if not filename: return
        
        self.form=ProgressBarForm("导出mgt文件中",4)
        th=Thread(ThreadStart(self.pbar_thread))
        th.IsBackground=True
        th.Start()
        
        #form=ProgressBarForm("导出mgt文件中")
        
        #open a new file
        file = open( filename, "w")
        
        #document data中的表格
        front_table_count=0
        if rs.GetDocumentData("mgt_info","front_table_count"):
            front_table_count=int(rs.GetDocumentData("mgt_info","front_table_count"))
            
        for i in range(front_table_count):
            nm="mgt_docu_table"+str(i+1)
            if rs.GetDocumentData(nm):
                file.write(rs.GetDocumentData("mgt_docu_table_names",nm))
                for entry in rs.GetDocumentData(nm):
                    file.write(rs.GetDocumentData(nm,entry))
                file.write("\n")
        self.form.Invoke(self.form.invoker)
        #form.bar_adjust(10)
        
        # Nodes
        file.write("*NODE    ; Nodes\n")
        file.write("; iNO, X, Y, Z\n")
        for coord in self.m_joints_dict:
            joint=self.m_joints_dict[coord]
            file.write("     %s, %s, %s, %s\n"%(joint.m_number,joint.m_X,joint.m_Y,joint.m_Z))
        file.write("\n")
        self.form.Invoke(self.form.invoker)
        #form.bar_adjust(20)
        
        # Elements
        file.write("*ELEMENT    ; Elements\n")
        file.write("; iEL, TYPE, iMAT, iPRO, iN1, iN2, ANGLE, iSUB, EXVAL, iOPT(EXVAL2) ; Frame  Element\n")
        file.write("; iEL, TYPE, iMAT, iPRO, iN1, iN2, ANGLE, iSUB, EXVAL, EXVAL2, bLMT ; Comp/Tens Truss\n")
        file.write("; iEL, TYPE, iMAT, iPRO, iN1, iN2, iN3, iN4, iSUB, iWID , LCAXIS    ; Planar Element\n")
        file.write("; iEL, TYPE, iMAT, iPRO, iN1, iN2, iN3, iN4, iN5, iN6, iN7, iN8     ; Solid  Element\n")
                
        max_section_number=0
        max_thickness_number=0
        new_section_dict={}  #key为layer name，value为ipro编号
        new_thickness_dict={}
        if rs.GetDocumentData("mgt_info","max_section_number"):
            max_section_number=int(rs.GetDocumentData("mgt_info","max_section_number"))
        if rs.GetDocumentData("mgt_info","max_thickness_number"):
            max_thickness_number=int(rs.GetDocumentData("mgt_info","max_thickness_number"))
            
        for frame in self.m_frames:
            if rs.GetUserText(frame.m_GUID,"mgt_element") is not None:
                str_tmp=rs.GetUserText(frame.m_GUID,"mgt_element")
                ls=str_tmp.split(",")
                if len(ls)>=6:
                    ls[0]="  %s"%frame.m_number
                    ls[4]="   %s"%frame.m_start_joint.m_number
                    ls[5]="   %s"%frame.m_end_joint.m_number
                    ls[1]=" TRUSS "
                    if rs.GetUserText(frame.m_GUID,"mgt_type") is not None:
                        ls[1]=rs.GetUserText(frame.m_GUID,"mgt_type")
                    
                    layer_name=rs.ObjectLayer(frame.m_GUID)
                    if re.search("mgt_section",layer_name) is not None:
                        tmp_layer_n=int(layer_name.replace("mgt_section",""))
                        ls[3]="   %s"%(tmp_layer_n)
                        #max_section_number=max(tmp_layer_n,max_section_number)
                    else:
                        if new_section_dict.has_key(layer_name):
                            ls[3]="   %s"%(new_section_dict[layer_name])
                        else:
                            max_section_number+=1
                            ls[3]="   %s"%max_section_number
                            new_section_dict[layer_name]=max_section_number
                for i in range(len(ls)):
                    file.write(ls[i])
                    if(i!=len(ls)-1):
                        file.write(",")
            else:
                layer_name=rs.ObjectLayer(frame.m_GUID)
                if re.search("mgt_section",layer_name) is not None:
                    tmp_layer_n=int(layer_name.replace("mgt_section",""))
                    #max_section_number=max(tmp_layer_n,max_section_number)
                else:
                    if new_section_dict.has_key(layer_name):
                        tmp_layer_n=new_section_dict[layer_name]
                    else:
                        max_section_number+=1
                        tmp_layer_n=max_section_number
                        new_section_dict[layer_name]=tmp_layer_n
                file.write("  %s, TRUSS ,    1,     %s,   %s,   %s,     0,     0\n"%(frame.m_number,tmp_layer_n,frame.m_start_joint.m_number,frame.m_end_joint.m_number))
            #file.write("\n")
        #frame_count=len(self.m_frames)
        
        for area in self.m_areas:
            if rs.GetUserText(area.m_GUID,"mgt_element") is not None:
                str_tmp=rs.GetUserText(area.m_GUID,"mgt_element")
            else:
                str_tmp="  1471, PLATE ,    1,     1,   298,   278,   304,     0,     1,     0\n"
            
            ls=str_tmp.split(",")
            if (len(ls)>=8) & (len(area.m_joints)>=3):
                ls[0]="  %s"%(area.m_number)
                if len(area.m_joints)==3:
                    #根据mgt输出要求的顺序输出
                    ls[4]="   %s"%area.m_joints[0].m_number
                    ls[5]="   %s"%area.m_joints[2].m_number
                    ls[6]="   %s"%area.m_joints[1].m_number
                    ls[7]="   0"
                else:
                    #根据mgt输出要求的顺序输出
                    ls[4]="   %s"%area.m_joints[0].m_number
                    ls[5]="   %s"%area.m_joints[2].m_number
                    ls[6]="   %s"%area.m_joints[3].m_number
                    ls[7]="   %s"%area.m_joints[1].m_number
                ls[1]=" PLATE "
                if rs.GetUserText(area.m_GUID,"mgt_type") is not None:
                    ls[1]=rs.GetUserText(area.m_GUID,"mgt_type")
                
                layer_name=rs.ObjectLayer(area.m_GUID)
                if re.search("mgt_thickness",layer_name,) is not None:
                    tmp_layer_n=int(layer_name.replace("mgt_thickness",""))
                    ls[3]="   %s"%(tmp_layer_n)
                else:
                    if new_thickness_dict.has_key(layer_name):
                        ls[3]="   %s"%(new_thickness_dict[layer_name])
                    else:
                        max_thickness_number+=1
                        ls[3]="   %s"%max_thickness_number
                        new_thickness_dict[layer_name]=max_thickness_number
                            
            for i in range(len(ls)):
                file.write(ls[i])
                if(i!=len(ls)-1):
                    file.write(",")
            #file.write("\n")
        file.write("\n")
        #form.bar_adjust(40)
        self.form.Invoke(self.form.invoker)
        #document data中的表格
        for nm in self.m_mgt_docu_table_names:
            if int(nm.replace("mgt_docu_table",""))-1 in range(front_table_count):
            #if nm=="mgt_docu_table1" or nm=="mgt_docu_table2" or nm=="mgt_docu_table3" or nm=="mgt_docu_table4":
                continue
            if rs.GetDocumentData(nm):
                file.write(rs.GetDocumentData("mgt_docu_table_names",nm))
                for entry in rs.GetDocumentData(nm):
                    file.write(rs.GetDocumentData(nm,entry))
                file.write("\n")
        #form.bar_adjust(60)
        self.form.Invoke(self.form.invoker)
        # EndData
        file.write("*ENDDATA\n")
    
        # Close File
        file.close()
        self.form.Close()
        print("导出mgt文件完成！")
        
    def import_mgt(self):
        #prompt the user for a file to import
        filter = "mgt file (*.mgt)|*.mgt|All Files (*.*)|*.*||"
        filename = rs.OpenFileName("Open MGT File", filter)
        if not filename: return
        
        self.form=ProgressBarForm("导入mgt文件中",4)
        th=Thread(ThreadStart(self.pbar_thread))
        th.IsBackground=True
        th.Start()
        
        #read each line from the file
        file = open(filename, "r")
        #type=chardet.detect(file.read())
        #print(type)
        table=""
        joints_dict={} #以编号为key，生成point对象后以point的GUID为value
        frames=[] #中间存储形式,TRUSS BEAM TENSTR COMPTR都存
        frames_dict={}#以编号为key，以GUID为value，TRUSS BEAM TENSTR COMPTR都存
        planars=[]
        planars_dict={}
        line_index=0
        table_index=0  #从1开始
        front_table_index=0 #在*Node之前的表的数量
        max_node_number=0
        for line in file:
            #去掉“;”后的注释部分
            ls=line.split(";")
            str_tmp=ls[0]
                
            #如果当前行是表名行，判断是哪个表名行
            if re.search("\\*",str_tmp) is not None:
                table=""
                if re.search("\\*NODE",str_tmp) is not None:
                    table="NODE"
                    front_table_index=table_index
                elif re.search("\\*ELEMENT",str_tmp) is not None:
                    table="ELEMENT"
                elif re.search("\\*ENDDATA",str_tmp) is not None:
                    break
                #elif re.search("\\*K-FACTOR",str_tmp) is not None:
                    #table=""
                #elif re.search("\\*CONSTRAINT",str_tmp) is not None:
                    #table=""
                #elif re.search("\\*FLOORLOAD",str_tmp) is not None:
                    #table=""
                #elif re.search("\\*REBAR-MATL-CODE",str_tmp) is not None:
                    #table=""
                else:
                    table=line
                    table_index+=1
                    line_index=0
                continue
                
            #如果当前行不是表名行
            #生成joints_dict和frames
            #将document data直接存入
            max_node_number=0
            if table=="":
                continue
            elif table=="NODE":
                str_tmp.replace(" ","")
                ls=str_tmp.split(",")
                if len(ls)>=4:
                    obj_pt=rs.AddPoint((float(ls[1]),float(ls[2]),float(ls[3])))
                    joints_dict[int(ls[0])]=obj_pt
                    rs.SetUserText(obj_pt,"mgt_original_number",ls[0])
                    max_node_number=max(int(ls[0]),max_node_number)
            elif table=="ELEMENT":
                ls=str_tmp.split(",")
                if (len(ls)>=6) and (ls[1].replace(" ","")=="TRUSS" or ls[1].replace(" ","")=="BEAM" or ls[1].replace(" ","")=="TENSTR" or ls[1].replace(" ","")=="COMPTR"):
                    frames.append(line)
                elif (len(ls)>=8) and (ls[1].replace(" ","")=="PLATE" or ls[1].replace(" ","")=="PLSTRS" or ls[1].replace(" ","")=="PLSTRN" or ls[1].replace(" ","")=="AXISYM" or ls[1].replace(" ","")=="WALL"):
                    planars.append(line)
            else:
                line_index+=1
                rs.SetDocumentData("mgt_docu_table_names","mgt_docu_table"+str(table_index),table)
                rs.SetDocumentData("mgt_docu_table"+str(table_index),str(line_index),line)
        
        #将max_node_number存入Document Data
        rs.SetDocumentData("mgt_Node_original_number","max_number",str(max_node_number))
        #将*Node表格前的表格数量存入Document Data
        rs.SetDocumentData("mgt_info","front_table_count",str(front_table_index))
        
        file.close()
        
        self.form.Invoke(self.form.invoker)
        #form.pbar.PerformStep()

        #生成frames_dict和planars_dict并在rhino中生成模型
        max_section_number=0
        max_thickness_number=0
        max_element_number=0
        for text in frames:
            ls=text.split(",")
            if len(ls)<6: continue
            if joints_dict.has_key(int(ls[4])) and joints_dict.has_key(int(ls[5])):
                obj=rs.AddLine(joints_dict[int(ls[4])],joints_dict[int(ls[5])])
                layer_name=rs.AddLayer("mgt_section%d"%int(ls[3]))
                rs.ObjectLayer(obj,layer_name)
                max_section_number=max(int(ls[3]),max_section_number)
                rs.SetUserText(obj,"mgt_element",text)
                rs.SetUserText(obj,"mgt_type",ls[1])#ls[1]带空格
                rs.SetUserText(obj,"mgt_original_number",ls[0])
                #将节点对应点对象的GUID存入杆件对象的UserText中
                rs.SetUserText(obj,"start_point_guid",joints_dict[int(ls[4])])
                rs.SetUserText(obj,"end_point_guid",joints_dict[int(ls[5])])
                #frames_dict[int(ls[0])]=obj
                max_element_number=max(int(ls[0]),max_element_number)
        #form.pbar.PerformStep()
        self.form.Invoke(self.form.invoker)
        
        for text in planars:
            ls=text.split(",")
            if len(ls)<8: continue
            if joints_dict.has_key(int(ls[4])) and joints_dict.has_key(int(ls[5])) and joints_dict.has_key(int(ls[6])):
                tmp_pts=[]
                tmp_pts.append(joints_dict[int(ls[4])])
                tmp_pts.append(joints_dict[int(ls[5])])
                tmp_pts.append(joints_dict[int(ls[6])])
                #将面对象的节点对应所有点对象的GUID存成一个字符串，以“，”分隔，顺序为pt1-pt3-pt2
                tmp_pts_guids=str(joints_dict[int(ls[4])])+","+str(joints_dict[int(ls[6])])+","+str(joints_dict[int(ls[5])])
                
                if int(ls[7])!=0 & joints_dict.has_key(int(ls[7])):
                    tmp_pts.append(joints_dict[int(ls[7])])
                    #若有4个点，顺序为pt1-pt4-pt2-pt3
                    tmp_pts_guids=str(joints_dict[int(ls[4])])+","+str(joints_dict[int(ls[7])])+","+str(joints_dict[int(ls[5])])+","+str(joints_dict[int(ls[6])])
                    
                obj=rs.AddSrfPt(tmp_pts)
                layer_name=rs.AddLayer("mgt_thickness%d"%int(ls[3]))
                rs.ObjectLayer(obj,layer_name)
                max_thickness_number=max(int(ls[3]),max_thickness_number)
                rs.SetUserText(obj,"mgt_element",text)
                rs.SetUserText(obj,"mgt_type",ls[1])#ls[1]带空格
                rs.SetUserText(obj,"mgt_original_number",ls[0])
                rs.SetUserText(obj,"point_guids",tmp_pts_guids)#将面对象的节点对应所有点对象的GUID字符串存入面对象的UserText中
                #planars_dict[int(ls[0])]=obj
                max_element_number=max(int(ls[0]),max_element_number)
        #form.pbar.PerformStep()
        self.form.Invoke(self.form.invoker)

        rs.SetDocumentData("mgt_info","max_section_number",str(max_section_number))
        rs.SetDocumentData("mgt_info","max_thickness_number",str(max_section_number))
        rs.SetDocumentData("mgt_Element_original_number","max_number",str(max_element_number))
        self.form.Close()
        print("从mgt文件导入完成！")
    
    def get_max_min_angle(self):
        if (self.m_joints_dict is None) or (self.m_frames is None):return
        
        max_angle=0
        max_joint=self.m_frames[0].m_start_joint
        max_frame1=self.m_frames[0]
        max_frame2=self.m_frames[0]
        min_angle=180
        min_joint=self.m_frames[0].m_start_joint
        min_frame1=self.m_frames[0]
        min_frame2=self.m_frames[0]
        for coord in self.m_joints_dict:
            tmp_max_angle,tmp_max_f1,tmp_max_f2,tmp_min_angle,tmp_min_f1,tmp_min_f2=self.m_joints_dict[coord].cal_max_min_frame_angle()
            if tmp_max_angle>max_angle:
                max_angle=tmp_max_angle
                max_joint=self.m_joints_dict[coord]
                max_frame1=tmp_max_f1
                max_frame2=tmp_max_f2
            if tmp_min_angle<min_angle:
                min_angle=tmp_min_angle
                min_joint=self.m_joints_dict[coord]
                min_frame1=tmp_min_f1
                min_frame2=tmp_min_f2
        return max_angle,max_joint,max_frame1,max_frame2,min_angle,min_joint,min_frame1,min_frame2

    def get_max_frame_length(self):
        if (self.m_joints_dict is None) or (self.m_frames is None):return
        
        max_length=self.m_frames[0].m_Length
        longest_frame=self.m_frames[0]
        for frame in self.m_frames:
            if frame.m_Length>max_length:
                max_length=frame.m_Length
                longest_frame=frame
        return max_length,longest_frame

    def display_grid_info(self):
        self.set_data_s2k()
        max_angle,max_joint,max_frame1,max_frame2,min_angle,min_joint,min_frame1,min_frame2=self.get_max_min_angle()
        max_length,longest_frame=self.get_max_frame_length()
    
        form=DisplayGridInfoForm(max_angle,min_angle,max_length)
        if( form.ShowDialog() == DialogResult.OK ):
            rhinoscript.geometry.AddTextDot("最大杆件夹角=%s"%max_angle, (max_joint.m_X,max_joint.m_Y,max_joint.m_Z))
            rhinoscript.geometry.AddTextDot("最小杆件夹角=%s"%min_angle, (min_joint.m_X,min_joint.m_Y,min_joint.m_Z))
    
            
if( __name__ == "__main__" ):
    model=ModelInfo()
    
    #model.import_s2k()
    model.export_s2k()
    #model.import_mgt()
    #model.export_mgt()
    
    #model.display_grid_info()
    
    
    