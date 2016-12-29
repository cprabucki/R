# Varias gráficas compartiendo eje X (diferentes ejes Y)

library("grid")

P=c(rnorm(52, 3000, 1000))
Q=c(rnorm(52, 10, 200))
R=c(rnorm(52, 50000, 0.5))

plot1 <- qplot(1:length(P),R, geom="line", colour = I("red"))
plot2 <- qplot(1:length(Q),P, geom="line", colour = I("blue"))
plot3 <- qplot(1:length(R), Q, geom="line", colour = I("green"))

grid.newpage()
grid.draw(rbind(ggplotGrob(plot1), ggplotGrob(plot2), ggplotGrob(plot3), size="last"))