export function stampAngle(value){const n=Number(value)||0;return Math.round(((n+180)%360+360)%360-180);}
export function stampGeometry(pos,ratio,pageWidth,pageHeight){
  const width=pos.w*pageWidth,height=width/ratio,cx=pos.x*pageWidth+width/2,cy=pos.y*pageHeight+height/2;
  const angle=stampAngle(pos.angle),radians=angle*Math.PI/180,c=Math.cos(radians),s=Math.sin(radians);
  const point=(x,y)=>({x:cx+x*c-y*s,y:cy+x*s+y*c});
  return {width,height,cx,cy,angle,radians,corners:[point(-width/2,-height/2),point(width/2,-height/2),point(width/2,height/2),point(-width/2,height/2)],extentX:(Math.abs(width*c)+Math.abs(height*s))/2,extentY:(Math.abs(width*s)+Math.abs(height*c))/2};
}
export function constrainStamp(pos,ratio,pageWidth,pageHeight){
  let g=stampGeometry(pos,ratio,pageWidth,pageHeight);
  const scale=Math.min(1,pageWidth/(2*g.extentX),pageHeight/(2*g.extentY));
  const width=g.width*scale,height=g.height*scale,ex=g.extentX*scale,ey=g.extentY*scale;
  const cx=Math.max(ex,Math.min(pageWidth-ex,g.cx)),cy=Math.max(ey,Math.min(pageHeight-ey,g.cy));
  return {...pos,angle:g.angle,w:width/pageWidth,x:(cx-width/2)/pageWidth,y:(cy-height/2)/pageHeight};
}
export function resizeStamp(pos,ratio,pageWidth,pageHeight,dx,dy){
  const g=stampGeometry(pos,ratio,pageWidth,pageHeight),c=Math.cos(g.radians),s=Math.sin(g.radians);
  const vx=c-s/ratio,vy=s+c/ratio,delta=(dx*vx+dy*vy)/(vx*vx+vy*vy);
  const width=Math.max(.035*pageWidth,Math.min(.6*pageWidth,g.width+delta)),difference=width-g.width;
  const cx=g.cx+difference*vx/2,cy=g.cy+difference*vy/2;
  return constrainStamp({...pos,w:width/pageWidth,x:(cx-width/2)/pageWidth,y:(cy-width/ratio/2)/pageHeight},ratio,pageWidth,pageHeight);
}
