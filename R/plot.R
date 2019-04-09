#' Generate graphs for kelp forest monitoring
#'
#' @param inDF input data frame
#' @param inLevels input levels
#' @param prefix prefix for output figures
#'
#' @return NULL generates output figures as jpg
#' @export
#'
#' @examples
generate_graphs <- function(inDF, inLevels, prefix){

  # Only get a list of species that are from the most current year (such as 2017)
  inTable_year <- inDF %>%
    filter(SurveyYear == currentYear)

  inDF$IslandName <- factor(inDF$IslandName, levels=inLevels)

  species <- unique(inTable_year$AlternateTaxonName)
  i <- 1
  for (spp in species) {
    species_count <- length(species)

    print(glue("Generating graph for: {spp}, {i} of {species_count}"))

    print("Subsetting by species")
    # Subselect just one of the species from the list
    spp_sub <- inDF %>%
      filter( AlternateTaxonName == spp)

    # Get the rank for naming the file
    rank <- unique(spp_sub$Species)
    rank<- rank[1]

    theme_set(theme_cowplot(font_size=4)) # reduce default font size

    print("Making ggList")
    # Use a function to get a list of all the graphs needed to join to make one page of graphs
    ggList <- lapply(split(spp_sub, spp_sub$IslandName), function(i) {
      ggplot(i, aes(x = SurveyYear, y = MeanDensity_sqm, color = SiteCode, linetype = SiteCode)) +
        geom_point(size = 1, alpha = 0.5, show.legend = FALSE)+
        geom_line(size = 0.5, alpha = 0.5)+
        #facet_wrap(~IslandCode, scales = "free_y", ncol = 1) +
        #facet_wrap("LongOrder", scales = "free_y") +
        #facet_grid(rows = vars(IslandCode), scales = "free_y") +
        #geom_smooth(method = loess, size = 0.4) +

        ggtitle(glue("{spp}")) +
        #ylab('Percent Cover')+
        #ylab(expression("#/600 m"^"3"))+
        ylab(expression("Mean density/m"^"2"))+
        theme(legend.position = "right", legend.text = element_text(size = 1), legend.title = element_text(size = 9), legend.key.size = unit(1, 'cm')) +
        guides(color=guide_legend(ncol=1, keyheight = 0.5)) +
        theme_bw() +
        theme(panel.grid.minor.x = element_blank())+
        scale_x_continuous(breaks = seq(1980, 2020, by = 1),expand = c(0.01,0.01)) +
        scale_y_continuous(limits=c(0, max(spp_sub$MeanDensity_sqm) * 1.1)) +
        scale_colour_manual("Site Code", values = c('Red','green3','Blue','Black','Red','green3','Blue','Black','purple3','darkorange2')) +
        theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1), axis.title.x = element_blank()) +
        theme(plot.title = element_text(size = 10, hjust = 1))+
        scale_linetype_manual("Site Code", values = c('solid','solid','solid','solid','dashed','dashed','dashed','dashed','solid','solid'))
    })

    # This is totally screwing up the graph shifting everything down. Can't figure out why.
    # title <- ggdraw() +
    #   draw_label(glue("{spp}"), fontface='plain')

    # plot as grid in 1 columns
    print("Plotting with cowplot")
    cowplot::plot_grid(plotlist = ggList,
                       ncol = 1,
                       align = 'v',
                       labels =  levels(inDF$IslandName),
                       label_fontface = "plain",
                       label_size = 12)

    # This is needed because some species names have special characters in them
    # and windows wont save file names with those characters in the file name
    print("Renaming species")
    spp_rename <- gsub("(>|/)", "", spp)

    # Save the file to disk in the Graphs folder
    print("Saving with ggsave")
    ggsave(filename = glue("Graphs\\{prefix}_{rank}_{spp_rename}.jpg"),
           plot = last_plot(),
           width = 7.5,
           height = 10,
           units = "in")

    i <- i + 1
  }
}
