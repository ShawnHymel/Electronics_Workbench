=begin
#-----------------------------------------------------------------------
Copyright 2008 TIG
Permission to use, copy, modify, and distribute this software for 
any purpose and without fee is hereby granted, provided something the 
above copyright notice appear in all copies.
THIS SOFTWARE IS PROVIDED "AS IS" AND WITHOUT ANY EXPRESS OR
IMPLIED WARRANTIES, INCLUDING, WITHOUT LIMITATION, THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE.
#-----------------------------------------------------------------------
Name        : 	ComponentReporter++.rb
Type        : 	Tool
Description : 	Makes reports on components etc in a model...
                Based on ComponentReporter+
Menu Item   : 	Plugins >> Component++ Report
Usage       : 	Answer prompts - for using just selected instances and 
				mining for nested instances etc.  
				If model name=XXXX reports is written into files 
				XXXX-Component+Report.csv for Component Definitions,
				XXXX-Instances+Report.csv for Component Instances, 
				XXXX-Parentage+Report.csv gives Outliner type report.  
				If the report is already open [=error] you are told.
				The reports are made in the model's folder.  
				CSV files are readable by Excel & similar programs.
				Commas [,] in reported names etc are trapped into ';'.  
				Try to avoid naming components, groups, layers etc with 
				commas, IF you are going to use tools like this !
				Area format = sqm
Version     :	1.0 20080722	First Issue.
              1.1 20090409  Glitch on multiple materials fixed.
			  1.2 20121008  MAC compatibility improved.
#-----------------------------------------------------------------------
=end
require 'sketchup.rb'
### do class
class Reporter
### do def
 def Reporter::components_plus_plus
   model=Sketchup.active_model
### check model is saved...
   mpath=model.path
   if mpath==""
      UI.messagebox("This 'Untitled' new Model must be Saved\nbefore making Component+ Reports !\nExiting... ")
      return nil
   end
   mpath=File.dirname(mpath)#(mpath.split("\\")[0..-2]).join("/")###strip off file name
   mname=model.title
   entities=model.entities
   ss=model.selection
   ssents=ss.to_a
### start undo...
   model.start_operation("Component+ Reports")
### show VCB and status info...
   Sketchup::set_status_text(("COMPONENT+ REPORTER..." ), SB_PROMPT)
   Sketchup::set_status_text(" ", SB_VCB_LABEL)
   Sketchup::set_status_text(" ", SB_VCB_VALUE)
### setup cdefn list...
   clist=[]
   model.definitions.each{|c|
      if c.count_instances > 0 and not c.group?
        clist.push([c.name.tr(",",";"),c.count_instances.to_s,c.description.tr(",",";"),c.guid])
      end#if
   } 
   clist.sort!
### if no components exit...
   if not clist[0]
      UI.messagebox("There are NO Components in this Model to make Component+ Reports !\nExiting... ")
      return nil
   end#if
### write to csv file...
   ccsv=File.join(mpath, mname+"-Component+Report.csv")
   begin
     file=File.new(ccsv,"w")
   rescue### trap if open
     UI.messagebox("Component Report File:\n\n  "+ccsv+"\n\nCannot be written - it's probably already open.\nClose it and try making the Report again...\n\nExiting...")
	 return nil
   end
   file.puts("DEFN-NAME,COUNT,DESCRIPTION,GUID\n\n") 
   ### title ### add what you want to list here
   clist.each{|c|
     file.puts(c[0]+","+c[1]+","+c[2]+","+c[3])
   }
   file.close
### check you really want to do compos in just the selection...
### but must check if ss has any compo instances in it.
   compo_selected=false
   ss.each{|c|compo_selected=true if c.typename=="ComponentInstance"}
   sel=""
   if compo_selected
      if UI.messagebox("Do you want an Instance Report for JUST the Selection ?  ",MB_YESNO,"Just Selection ?")==6 ### 6=YES 7=NO
	    entities=ssents
	    sel=" (for Selection)"
	  end#if
   end#if
### ask if you want to mine instances
   miner=false
   miner=true if UI.messagebox("Do you want to 'Mine' for ALL Nested Instances ?  ",MB_YESNO,"Mine ?")==6 ### 6=YES 7=NO
### NOW make a list of all "instances" with details
### setup instances list...
   inlist=[]; all_list=[]
   entities.each{|c|
      inlist.push(c) if c.typename=="ComponentInstance"
	  all_list.push(c) if c.typename=="ComponentInstance" or c.typename=="Group"
   }
### now mine for instances
##########################
def Reporter::miner(ents)
  list=ents
  ents.each{|e|
    elist=[]
    if e.typename=="ComponentInstance"
	  e.definition.entities.each{|c|
        elist.push(c) if c.typename=="ComponentInstance" or c.typename=="Group"
	  }
	end#if
	if e.typename=="Group"
      e.entities.each{|c|
        elist.push(c) if c.typename=="ComponentInstance" or c.typename=="Group"
	  }
	end#if
	xlist=Reporter.miner(elist)
	list=list+xlist
  }#end each
  return list
end#def
##########################
   mine=""
   if miner
      mine=" (Mined)"
      full_list=Reporter.miner(all_list)
	  inlist=[]
	  full_list.each{|c|
	    inlist.push(c) if c.typename=="ComponentInstance"
	  }### misses out Groups at end
   end#if
   ilist=[]
   inlist.each{|c|
     dname="'"+c.definition.name.tr(",",";")+"'"
	 ###''avoids loss of front zeros
	 ###tr , >> ; avoids names etc with , in messing up csv
	 iname="'"+c.name.tr(",",";")+"'"
	 lname="'"+c.layer.name.tr(",",";")+"'"
	 mat=c.material
	 if mat
	   cname="'"+mat.name.tr(",",";")+"'"
	 else
	   cname="'<Default>'"
	 end#if
	 area=0.0
	 c.definition.entities.each{|e|area=area+e.area if e.typename=="Face"}
	 area=0.09290304*area/144 ### sq m
   area=area.to_s.tr(",",".")### trap for comma as decimal point...
   areas=["aa",123,"bb",124,"cc",234,"dd",333]
	 id=c.entityID.to_s
	 x=(c.bounds.max.x-c.bounds.min.x).to_l.to_s
	 y=(c.bounds.max.y-c.bounds.min.y).to_l.to_s
	 z=(c.bounds.max.z-c.bounds.min.z).to_l.to_s
	 parent=c.parent
	 if parent==model
	   pname="'"+mname.tr(",",";")+"'<Model>"
	 else
	   pname="'"+parent.name.tr(",",";")+"'"
	   pname=pname+"<Group>" if parent.group?
	   pname=pname+"<Component>" if not parent.group?
	 end#if
     ilist.push([dname,iname,lname,cname,id,x,y,z,pname,area,areas].flatten!)
   }
   ilist.sort!
### write to csv file...
   icsv=mpath+"/"+mname+"-Instances+Report.csv"
   begin
     file=File.new(icsv,"w")
   rescue### trap if open
     UI.messagebox("Instances Report File:\n\n  "+icsv+"\n\nCannot be written - it's probably already open.\nClose it and try making the Report again...\n\nExiting...")
	 return nil
   end
   #file=File.new(icsv,"w")
   file.puts("DEFN-NAME,PARENT[NAME<TYPE>],INST-NAME,LAYER,MATERIAL,ID,X,Y,Z,AREA,SUB-MAT-AREAS\n\n")if miner
   file.puts("DEFN-NAME,INST-NAME,LAYER,MATERIAL,ID,X,Y,Z,AREA,SUB-MAT-AREAS\n\n")if not miner
  ilist.each{|c|
   ###################????????????????????????
   xx=""; c[10..-1].each{|x|;xx=","+x.to_s}
   if miner
	  file.puts(c[0]+","+c[8]+","+c[1]+","+c[2]+","+c[3]+","+c[4]+","+c[5]+","+c[6]+","+c[7]+","+c[9]+xx)
	 else
	  file.puts(c[0]+","+c[1]+","+c[2]+","+c[3]+","+c[4]+","+c[5]+","+c[6]+","+c[7]+","+c[9]+xx)
	 end#if
   }
   file.close
### make parentage report (outliner mimic)
   enlist=[]
   entities.each{|e|
     enlist.push(e) if e.typename=="ComponentInstance" or e.typename=="Group"
   }
   elist=Reporter.miner(enlist)
   relist=elist
   elist.each{|e|
     relist=relist-[e.parent]
   }
   plist=[]
   elist.each{|e|
	 ee=e
	 parent=ee.parent
	 tree=[ee]
	 while parent != model
	   pa=ee.parent
	   tree.unshift(pa) if pa != model
	   ee=pa
	   parent=ee.parent
	 end#while
	 txt="'"+mname.tr(",",";")+"'<Model>,"
	 name=""
	 tree.each{|n|
	   name="'"+n.name.tr(",",";")+"'"+n.definition.name.tr(",",";")+"<Instance>," if n.typename=="ComponentInstance"
	   name="'"+n.name.tr(",",";")+"'<Definition>," if n.typename=="ComponentDefinition" and not n.group?
	   name="'"+n.instances[0].name.tr(",",";")+"'<Group>," if n.typename=="ComponentDefinition" and n.group?
	   name="'"+n.name.tr(",",";")+"'<Group>," if n.typename=="Group"
	   txt=txt+name
     }
	 txt=txt+"\n"
	 plist.push(txt)
   }
   plist.sort!
### write to csv file...
   pcsv=mpath+"/"+mname+"-Parentage+Report.csv"
   begin
     file=File.new(pcsv,"w")
   rescue### trap if open
     UI.messagebox("Parentage Report File:\n\n  "+pcsv+"\n\nCannot be written - it's probably already open.\nClose it and try making the Report again...\n\nExiting...")
	 return nil
   end
   plist.each{|e|file.puts(e)}
   file.close
#################
### say it's done...
   UI.messagebox("Component+ Report written into:\n\n"+ccsv+"\n\nInstance+ Report"+sel+mine+" written into:\n\n"+icsv+"\n\n"+"Parentage+ Report"+sel+" written into:\n\n"+pcsv+"\n\n")
### commit undo...
   model.commit_operation
 end#def components
end#class
###########
### do menu
if not file_loaded?(File.basename(__FILE__))
    add_separator_to_menu("Plugins")
    UI.menu("Plugins").add_item("Component++ Report"){Reporter.components_plus_plus}
end#if
file_loaded(File.basename(__FILE__))
###
