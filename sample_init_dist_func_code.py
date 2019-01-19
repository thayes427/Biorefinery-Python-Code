#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Jan 18 19:14:02 2019

@author: Tom1
"""

def plot_init_dist(self):
        '''
        This function will plot the distribution of variable calls prior to running
        the simulation. This will enable users to see whether the distributions are as they expected.
        
        '''
        
        self.get_distributions()        
            
            
        if not self.graphs_created:
            
            if self.univar_row_num != 0: # basically, is this for univariate analysis
                row_num = 17
            else:
                row_num = 10
            
            fig_list =[]
            self.plot_list = []
            for var, values in self.simulation_dist.items():
                fig = Figure(figsize = (3,3), facecolor=[240/255,240/255,237/255], tight_layout=True)
                a = fig.add_subplot(111)
                num_bins = 15
                a.hist(values, num_bins, facecolor='blue', edgecolor='black', alpha=1.0)
                a.set_title(var)
                fig_list.append(fig)
                self.plot_list.append(a)
            frame_canvas = Frame(self.current_tab)
            frame_canvas.grid(row=row_num, column=1, columnspan = 3,pady=(5, 0))
            frame_canvas.grid_rowconfigure(0, weight=1)
            frame_canvas.grid_columnconfigure(0, weight=1)
            frame_canvas.config(height = '10c', width='16c')
            
            main_canvas = Canvas(frame_canvas)
            main_canvas.grid(row=0, column=0, sticky="news")
            main_canvas.config(height = '10c', width='16c')
            
            vsb = Scrollbar(frame_canvas, orient="vertical", command=main_canvas.yview)
            vsb.grid(row=0, column=2,sticky = 'ns')
            main_canvas.configure(yscrollcommand=vsb.set)
            
            figure_frame = Frame(main_canvas)
            main_canvas.create_window((0, 0), window=figure_frame, anchor='nw')
            figure_frame.config(height = '10c', width='16c')
            frame_canvas.config(width='16c', height='10c')
            
            row_num = 0
            column = False
            self.graphs_currently_displayed = []
            for figs in fig_list:
                figure_canvas = FigureCanvasTkAgg(figs, master=figure_frame)
                self.graphs_currently_displayed.append(figure_canvas)
                if column:
                    col = 4
                else:
                    col = 1
                figure_canvas.get_tk_widget().grid(
                        row=row_num, column=col,columnspan=2, rowspan = 5, pady = 5,padx = 8, sticky=E)
    
                if column:
                    row_num += 5
                column = not column
            figure_frame.update_idletasks()
            main_canvas.config(scrollregion=figure_frame.bbox("all"))
    
        
        else:
            for f in self.plot_list:
                f.cla()
                f.clear()
            counter = 0
            num_bins = 15
            for var, values in self.simulation_dist.items():
                self.plot_list[counter].hist(values, num_bins, facecolor='blue', edgecolor='black', alpha=1.0)
                self.plot_list[counter].set_title(var)
                counter += 1
            for fig in self.graphs_currently_displayed:
                fig.draw()
                #fig.get_tk_widget().destroy()
                #fig.get_tk_widget().forget()
                #self.main_canvas.delete(ALL)
            
        
        self.graphs_created = True