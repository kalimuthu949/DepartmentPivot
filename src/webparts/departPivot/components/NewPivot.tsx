import * as React from 'react';
import TreeView from '@material-ui/lab/TreeView';
import ExpandMoreIcon from '@material-ui/icons/ExpandMore';
import ChevronRightIcon from '@material-ui/icons/ChevronRight';
import TreeItem from '@material-ui/lab/TreeItem';
import { graph } from "@pnp/graph/presets/all";
import "../assets/Css/NewPivot.css"
import { MSGraphClient } from "@microsoft/sp-http";
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
let allUserData=[];
export default function NewPivot(props) 
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


    function bindData(data)
    {
      const users = [];
      let depts=[];
      for (let i = 0; i < data.length; i++) 
      {
        let userIdentity=data[i].identities[0].issuer;
        let userPrinName=data[i].userPrincipalName?data[i].userPrincipalName:"";
        if(!props.propertyToggle)
        {
          if(userIdentity.toLowerCase()=="hosthealthcare.onmicrosoft.com" && !userPrinName.includes('#EXT#'))
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
          }
        }
        else
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
        }
        
      }

      depts=removeDuplicates(depts);
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
                        designations[i].UserCount=designations[i].UserCount+1;
                        designations[i].Designations[0].count=designations[i].Designations[0].count+1;
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

    let sortedArray = designations.sort(function(a, b){
        if(a.Dept < b.Dept) { return -1; }
        if(a.Dept > b.Dept) { return 1; }
        return 0;
    });

      console.log(sortedArray);
      setdesignationdetails([...sortedArray]);
      setdepartment([...depts]);
      setPeopleList([...users]);
      setloader(false);
    }

    async function getNextusers(skiptoken) 
    {
      //await graph.users.select("department,mail,id,displayName,jobTitle,mobilePhone").top(10)
      await props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient) => {
        client
          .api("users")
          .select(
            "department,mail,id,displayName,jobTitle,mobilePhone,identities,userPrincipalName"
          )
          .top(999)
          .skipToken(skiptoken)
          .get().then(function (data)
          {
            allUserData=[...allUserData,...data.value];        
            let strtoken = "";
                if (data["@odata.nextLink"]) 
                {
                  strtoken = data["@odata.nextLink"].split("skipToken=")[1];
                  getNextusers(data["@odata.nextLink"].split("skipToken=")[1]);
                }
                else
                {
                  bindData(allUserData);
                }
          
              }).catch(function (error) {
                console.log(error)
                setloader(false);
              })
            });
    }

  async function getallusers() {
    allUserData=[];
    //await graph.users.select("department,mail,id,displayName,jobTitle,mobilePhone").top(10)
    await props.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient) => {
        client
          .api("users")
          .select(
            "department,mail,id,displayName,jobTitle,mobilePhone,identities,userPrincipalName"
          )
          .top(999)
          .get().then(function (data) 
          {
            allUserData=data.value;        
            let strtoken = "";
              if (data["@odata.nextLink"]) 
              {
                strtoken = data["@odata.nextLink"].split('skiptoken=')[1];
                getNextusers(data["@odata.nextLink"].split('skiptoken=')[1]);
              }
              else
              {
                bindData(data.value);
              }
        
            }).catch(function (error) {
              console.log(error)
              setloader(false);
            })
          });
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
        let sortedArraydesignation = item.Designations.sort(function(a, b){
          if(a.Designation < b.Designation) { return -1; }
          if(a.Designation > b.Designation) { return 1; }
          return 0;
      });
        return(<TreeItem nodeId={index.toString()} label={item.Dept +" ("+item.UserCount+")"}>
            {sortedArraydesignation.map(function(item,index)
            {
                let labelvalue=item.Designation+" ("+item.count+")";
                return(<TreeItem nodeId={designationdetails.length.toString()} label={labelvalue} />)
            })}
        </TreeItem>)
      })}
    

    </TreeView></div>
  );
}