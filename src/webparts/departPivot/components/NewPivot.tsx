import * as React from 'react';
import TreeView from '@material-ui/lab/TreeView';
import ExpandMoreIcon from '@material-ui/icons/ExpandMore';
import ChevronRightIcon from '@material-ui/icons/ChevronRight';
import TreeItem from '@material-ui/lab/TreeItem';
import { graph } from "@pnp/graph/presets/all";
import "../assets/Css/NewPivot.css"
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';

export default function NewPivot() 
{

    const [peopleList, setPeopleList] = React.useState([]);
    const [department,setdepartment]= React.useState([]);
    const [designationdetails,setdesignationdetails]= React.useState([]);
    const [loader,setloader]= React.useState(false);
  
    React.useEffect(function()
    {
      setloader(true);  
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
                designations.push({Dept:depts[i],UserCount:0,Designations:[]})
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
                                designations[index].UserCount=designations[index].UserCount+1;
                                designations[index].Designations[0].count=designations[index].Designations[0].count+1;
                            }
                            else
                            {
                              designations[i].UserCount=designations[i].UserCount+1;
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
              setloader(false);
        
            }).catch(function (error) {
              console.log(error)
              setloader(false);
            })
          }
  
    return (<div className='clsPivot'>
      {loader?<div className="spinnerBackground"><Spinner className="clsSpinner" size={SpinnerSize.large} /></div>:<></>}
    <TreeView
      aria-label="file system navigator"
      defaultCollapseIcon={<ExpandMoreIcon />}
      defaultExpandIcon={<ChevronRightIcon />}
    >
      {designationdetails.map(function(item,index)
      {
        let count=item.Designations.length;
        return(<TreeItem nodeId={index.toString()} label={item.Dept +" "+item.UserCount}>
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