export class LocalFileStore {
  constructor(){this.directory=null;this.handle=null;this.fileName='';this.stamp=null;this.writing=false;}
  async selectDirectory(){this.directory=await window.showDirectoryPicker({id:'sheet-workbench',mode:'readwrite'});this.handle=null;this.fileName='';this.stamp=null;return this.list();}
  async list(){if(!this.directory)return[];const files=[];for await(const [name,handle] of this.directory.entries())if(handle.kind==='file'&&/\.xlsx$/i.test(name)&&!name.startsWith('~$'))files.push({name,handle});return files.sort((a,b)=>a.name.localeCompare(b.name,'zh-CN'));}
  async open(handle){const file=await handle.getFile();const bytes=await file.arrayBuffer();return{bytes,file,handle};}
  associate(handle,file){this.handle=handle;this.fileName=handle.name;this.stamp={size:file.size,lastModified:file.lastModified};}
  async create(name,bytes){
    if(!this.directory)throw new Error('请先选择本地目录。');
    try{await this.directory.getFileHandle(name);throw new Error('目录中已有同名文件，请换一个名称。');}catch(e){if(e.name!=='NotFoundError')throw e;}
    const handle=await this.directory.getFileHandle(name,{create:true});let stream;
    try{stream=await handle.createWritable();await stream.write(bytes);await stream.close();}catch(e){try{await stream?.abort();}catch{}throw e;}
    this.associate(handle,await handle.getFile());return handle;
  }
  async save(bytes,{manual=false}={}){
    if(this.writing)throw new Error('正在保存，请稍候。');
    if(!this.handle)throw new Error('请先关联本地 Excel 文件。');
    this.writing=true;let stream;
    try{
      let permission=await this.handle.queryPermission({mode:'readwrite'});
      if(permission!=='granted'&&manual)permission=await this.handle.requestPermission({mode:'readwrite'});
      if(permission!=='granted')throw new Error('文件写入权限已失效，请点击手动保存重新授权。');
      const file=await this.handle.getFile();
      if(this.stamp&&(file.size!==this.stamp.size||file.lastModified!==this.stamp.lastModified)){const e=new Error('本地文件已被其他软件修改，当前修改尚未保存。');e.code='CONFLICT';throw e;}
      stream=await this.handle.createWritable();await stream.write(bytes);await stream.close();this.associate(this.handle,await this.handle.getFile());
    }catch(e){try{await stream?.abort();}catch{}throw e;}finally{this.writing=false;}
  }
}
