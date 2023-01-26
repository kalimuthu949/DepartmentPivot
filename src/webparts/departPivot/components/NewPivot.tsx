import * as React from 'react';
import TreeView from '@material-ui/lab/TreeView';
import ExpandMoreIcon from '@material-ui/icons/ExpandMore';
import ChevronRightIcon from '@material-ui/icons/ChevronRight';
import TreeItem from '@material-ui/lab/TreeItem';
import { graph } from "@pnp/graph/presets/all";
import "../assets/Css/NewPivot.css"

export default function NewPivot() 
{

    const [peopleList, setPeopleList] = React.useState([]);
    const [department,setdepartment]= React.useState([]);
    const [designationdetails,setdesignationdetails]= React.useState([]);
  
    React.useEffect(function()
    {
        getallusers();
    },[])

    function removeDuplicates(arr) {
        return arr.filter((item, 
            index) => arr.indexOf(item) === index);
      }

    async function getallusers() {
            await graph.users.select("department,mail,id,displayName,jobTitle,mobilePhone").top(999).get().then(function (data) {
              console.log(data);
              const users = [];
              let depts=[];
              for (let i = 0; i < data.length; i++) 
              {
                users.push({
                  imageUrl: "/_layouts/15/userphoto.aspx?size=L&username=" + data[i].mail,
                  isValid: true,
                  Email: data[i].mail,
                  ID: data[i].id,
                  key: i,
                  text: data[i].displayName,
                  jobTitle:data[i].jobTitle,
                  mobilePhone:data[i].mobilePhone,
                  department:data[i].department
                })
        
                if(data[i].department)
                depts.push(data[i].department)
                
               depts=removeDuplicates(depts);
    
              }
              let designations=[];
              for(let i=0;i<depts.length;i++)
              {
                designations.push({Dept:depts[i],Designations:[]})
                for(let j=0;j<users.length;j++)
                {
                    if(users[j].department==depts[i])
                    {
                        if(users[j].jobTitle)
                        {
                            let obj = "";
                            if(designations[i].Designations.length>0)
                            {
                                if(users[j].jobTitle)
                                obj=designations[i].Designations.find(o => o.Designation == users[j].jobTitle);
                            }
                            if(obj)
                            {
  
                                let index = designations[i].Designations.findIndex(o => o.Designation == users[j].jobTitle);
                                designations[index].Designations[0].count=designations[index].Designations[0].count+1;
                            }
                            else
                            {
                                designations[i].Designations.push({Designation:users[j].jobTitle,count:1});
                            }
                        }
                        
                    }
                }
              }

              console.log(designations);
              setdesignationdetails([...designations]);
              setdepartment([...depts]);
              setPeopleList([...users]);
        
            }).catch(function (error) {
              console.log(error)
            })
          }
  
    return (<div className='clsPivot'>
    <TreeView
      aria-label="file system navigator"
      defaultCollapseIcon={<ExpandMoreIcon />}
      defaultExpandIcon={<ChevronRightIcon />}
    >
      {designationdetails.map(function(item,index){
        return(<TreeItem nodeId={index.toString()} label={item.Dept}>
            {item.Designations.map(function(item,index)
            {
                let labelvalue=item.Designation+" ("+item.count+")";
                return(<TreeItem nodeId={designationdetails.length.toString()} label={labelvalue} />)
            })}
        </TreeItem>)
      })}
    

    </TreeView></div>
  );
}