use std::{env, fs, path::PathBuf};
fn main() {
    if env::var("CARGO_CFG_TARGET_OS").as_deref() != Ok("windows") { return; }
    let out = PathBuf::from(env::var("OUT_DIR").unwrap());
    let path = out.join("excel-daily.ico");
    let mut dir = ico::IconDir::new(ico::ResourceType::Icon);
    for side in [16u32, 32, 48, 256] {
        let mut px = vec![0u8; (side * side * 4) as usize];
        for y in 0..side { for x in 0..side {
            let i = ((y * side + x) * 4) as usize;
            let panel = x > side/8 && x < side*7/8 && y > side/8 && y < side*7/8;
            let grid = panel && ((x > side/3 && x < side/3+side/18) || (y > side/2 && y < side/2+side/18));
            let (r,g,b,a) = if grid {(255,255,255,255)} else if panel {(26,130,72,255)} else {(0,0,0,0)};
            px[i..i+4].copy_from_slice(&[r,g,b,a]);
        }}
        let img = ico::IconImage::from_rgba_data(side, side, px);
        dir.add_entry(ico::IconDirEntry::encode(&img).unwrap());
    }
    let mut f = fs::File::create(&path).unwrap();
    dir.write(&mut f).unwrap();
    let mut res = winres::WindowsResource::new();
    res.set_icon(path.to_str().unwrap());
    res.set("ProductName", "Excel Daily");
    res.set("FileDescription", "MICS Excel Daily Hotkey Agent");
    res.compile().unwrap();
}
