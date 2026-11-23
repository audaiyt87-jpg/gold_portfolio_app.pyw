"""Native Tk Canvas live chart: no browser, JavaScript, or extra dependency."""
import math
from bisect import bisect_left
from datetime import datetime


def clean_points(samples, now, window, factor=1.0):
    points = []
    for row in samples:
        try:
            ts, value = float(row[0]), float(row[1]) * factor
            if math.isfinite(ts) and math.isfinite(value) and value > 0 and now-window <= ts <= now+1:
                points.append((ts, value))
        except (TypeError, ValueError, IndexError, OverflowError):
            continue
    points.sort()
    return points


def reduce_points(points, limit=600):
    """Preserve first/last and min/max in chronological buckets."""
    if len(points) <= limit:
        return list(points)
    result = [points[0]]
    bucket = max(1, math.ceil((len(points)-2) / ((limit-2)//2)))
    for start in range(1, len(points)-1, bucket):
        segment = points[start:min(start+bucket, len(points)-1)]
        lo = min(range(len(segment)), key=lambda i: segment[i][1])
        hi = max(range(len(segment)), key=lambda i: segment[i][1])
        result.extend(segment[i] for i in sorted(set((lo, hi))))
    result.append(points[-1])
    return result


def render_chart(canvas, points, now, window, theme, unit, privacy=False):
    canvas.delete('all')
    width, height = canvas.winfo_width(), canvas.winfo_height()
    if width < 220 or height < 150:
        return None
    muted = theme.get('muted', '#8d98a8')
    text = theme.get('text', '#edf2f7')
    border = theme.get('border', '#273244')
    if privacy:
        canvas.create_text(width/2, height/2, text='وضع الخصوصية — الرسم مخفي', fill=muted,
                           font=('Segoe UI', 11))
        return None
    left, right, top, bottom = 86, width-20, 35, height-38
    oldest, newest = now-window, now
    values = [p[1] for p in points]
    if not values:
        canvas.create_text(width/2, height/2, text='بانتظار أول تسعير حديث…', fill=muted,
                           font=('Segoe UI', 11))
        return None
    low, high = min(values), max(values)
    pad = max((high-low)*.15, abs(values[-1])*.00003, .003)
    low, high = low-pad, high+pad
    def x(ts):
        return left+(ts-oldest)/window*(right-left)
    def y(value):
        return top+(high-value)/(high-low)*(bottom-top)
    for index in range(5):
        value = low+(high-low)*index/4
        yy = y(value)
        canvas.create_line(left, yy, right, yy, fill=border)
        canvas.create_text(left-9, yy, text=f'{value:,.3f}', anchor='e', fill=muted,
                           font=('Segoe UI', 9))
    for index in range(4):
        ts = oldest+window*index/3
        anchor = 'w' if index==0 else 'e' if index==3 else 'center'
        fmt = '%m-%d %H:%M' if window > 86400 else '%H:%M:%S' if window <= 900 else '%H:%M'
        canvas.create_text(x(ts), bottom+18, text=datetime.fromtimestamp(ts).strftime(fmt),
                           anchor=anchor, fill=muted, font=('Segoe UI', 9))
    canvas.create_text(left, 14, text=unit, anchor='w', fill=text, font=('Segoe UI', 10, 'bold'))
    color = theme.get('profit', '#2ecc71') if values[-1] >= values[0] else theme.get('loss', '#ff5c5c')
    # Deliberately no smooth interpolation or invented intermediate samples.
    reduced = reduce_points(points)
    groups = [[]]
    for point in reduced:
        if groups[-1] and point[0]-groups[-1][-1][0] > max(20, window/200):
            groups.append([])
        groups[-1].append(point)
    for group in groups:
        if len(group)>1:
            coords = [coord for ts,v in group for coord in (x(ts),y(v))]
            canvas.create_line(*coords, fill=color, width=2.5, smooth=False)
        elif group:
            xx,yy=x(group[0][0]),y(group[0][1])
            canvas.create_oval(xx-2,yy-2,xx+2,yy+2,fill=color,outline='')
    last_x,last_y=x(points[-1][0]),y(values[-1])
    canvas.create_line(left,last_y,right,last_y,fill=color,dash=(4,4))
    canvas.create_oval(last_x-4,last_y-4,last_x+4,last_y+4,fill=color,outline='')
    canvas.create_text(right,14,text=f'{values[-1]:,.3f}',anchor='e',fill=color,
                       font=('Segoe UI',12,'bold'))
    return {'points':points,'xs':[x(p[0]) for p in points], 'x':x,'y':y,
            'bounds':(left,top,right,bottom),'color':color,'unit':unit}


def hover_chart(canvas, layout, event, theme):
    canvas.delete('hover')
    if not layout:
        return
    left,top,right,bottom=layout['bounds']
    if not left <= event.x <= right or not top <= event.y <= bottom:
        return
    xs=layout['xs']
    index=bisect_left(xs,event.x)
    candidates=[i for i in (index-1,index) if 0 <= i < len(xs)]
    if not candidates:
        return
    nearest=min(candidates,key=lambda i:abs(xs[i]-event.x))
    ts,value=layout['points'][nearest]
    xx,yy=layout['x'](ts),layout['y'](value)
    muted=theme.get('muted','#8d98a8')
    canvas.create_line(xx,top,xx,bottom,fill=muted,dash=(3,3),tags='hover')
    canvas.create_oval(xx-5,yy-5,xx+5,yy+5,outline=layout['color'],width=2,tags='hover')
    message=f'{datetime.fromtimestamp(ts).strftime("%H:%M:%S")}  |  {value:,.3f}'
    tip=canvas.create_text((left+right)/2,top+12,text=message,fill=theme.get('text','#edf2f7'),
                           font=('Segoe UI',10,'bold'),tags='hover')
    box=canvas.bbox(tip)
    if box:
        bg=canvas.create_rectangle(box[0]-6,box[1]-3,box[2]+6,box[3]+3,
                                   fill=theme.get('card','#12161e'),outline=muted,tags='hover')
        canvas.tag_lower(bg,tip)
