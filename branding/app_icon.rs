// Shared Windows icon for Excel Daily Agent, Classic Settings and Tauri Settings.
// Render at 4x subpixel sampling to keep the calendar and clock crisp at 16px.
use std::{fs::File, path::Path};
pub const SIZES: [u32; 7] = [16, 20, 24, 32, 48, 64, 256];
fn rounded(x:f32,y:f32,l:f32,t:f32,r:f32,b:f32,radius:f32)->bool {
    if x<l || x>r || y<t || y>b {return false}
    let dx=x-x.clamp(l+radius,r-radius);
    let dy=y-y.clamp(t+radius,b-radius);
    dx*dx+dy*dy<=radius*radius
}
fn segment(x:f32,y:f32,ax:f32,ay:f32,bx:f32,by:f32,width:f32)->bool{
    let vx=bx-ax;let vy=by-ay;
    let q=(((x-ax)*vx+(y-ay)*vy)/(vx*vx+vy*vy)).clamp(0.0,1.0);
    let dx=x-ax-q*vx;let dy=y-ay-q*vy;
    dx*dx+dy*dy<=width*width/4.0
}
fn sample(x:f32,y:f32)->[u8;4]{
    if !rounded(x,y,2.0,2.0,62.0,62.0,13.0){return [0,0,0,0]}
    let k=((x+y-4.0)/120.0).clamp(0.0,1.0);
    let mut c=[(31.0-17.0*k) as u8,(143.0-66.0*k) as u8,(92.0-37.0*k) as u8,255];
    // Calendar sheet and dark header separator.
    if rounded(x,y,11.0,14.0,49.0,51.0,4.0){c=[246,255,249,255]}
    for cx in [21.0,39.0] {
        if rounded(x,y,cx-2.5,9.0,cx+2.5,20.0,2.0){c=[201,244,218,255]}
    }
    if x>=11.0&&x<=49.0&&y>=23.0&&y<=26.0{c=[29,121,76,255]}
    // Two clear rows of schedule blocks.
    for yy in [31.0,40.0] {for xx in [17.0,28.0,39.0]{
        if rounded(x,y,xx,yy,xx+6.0,yy+6.0,1.0){c=[65,155,107,255]}
    }}
    // A contrasting clock overlaps the lower right, with a green outline.
    let d=(x-46.0)*(x-46.0)+(y-46.0)*(y-46.0);
    if d<=15.0*15.0{c=[16,78,51,255]}
    if d<=12.3*12.3{c=[255,208,106,255]}
    if segment(x,y,46.0,39.0,46.0,46.0,3.0)||
       segment(x,y,46.0,46.0,51.5,49.0,3.0){c=[22,66,47,255]}
    c
}
pub fn rgba(side:u32)->Vec<u8>{
    let mut pixels=vec![0; (side*side*4) as usize];
    for y in 0..side {for x in 0..side {
        let mut sum=[0u32;3];let mut count=0u32;
        for sy in 0..4 {for sx in 0..4 {
            let c=sample((x as f32+(sx as f32+0.5)/4.0)*64.0/side as f32,
                         (y as f32+(sy as f32+0.5)/4.0)*64.0/side as f32);
            if c[3]!=0 {count+=1;for channel in 0..3{sum[channel]+=c[channel] as u32}}
        }}
        let i=((y*side+x)*4) as usize;
        if count>0{for channel in 0..3{pixels[i+channel]=(sum[channel]/count) as u8}
            pixels[i+3]=(255*count/16) as u8;}
    }}
    pixels
}
pub fn write_icon(path:&Path){
    let mut dir=ico::IconDir::new(ico::ResourceType::Icon);
    for side in SIZES {
        let pixels=rgba(side);
        assert_eq!(pixels.len(),(side*side*4) as usize);
        assert_eq!(pixels[3],0,"icon corner must be transparent");
        assert!(pixels.chunks_exact(4).any(|p|p[3]==255),"icon must contain opaque artwork");
        let img=ico::IconImage::from_rgba_data(side,side,pixels);
        dir.add_entry(ico::IconDirEntry::encode(&img).expect("encode app icon"));
    }
    dir.write(&mut File::create(path).expect("create app icon")).expect("write app icon");
    let decoded=ico::IconDir::read(File::open(path).expect("reopen app icon")).expect("validate generated icon");
    assert_eq!(decoded.entries().len(),SIZES.len());
    for (entry,side) in decoded.entries().iter().zip(SIZES) {
        assert_eq!(entry.width(),side);assert_eq!(entry.height(),side);
        entry.decode().expect("decode generated icon image");
    }
}
